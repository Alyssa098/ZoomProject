Option Strict On
Option Explicit On

Imports System.Collections.Generic
Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.SqlTypes
Imports System.Linq
Imports System.Text
Imports System.Diagnostics
Imports System.Collections
Imports BT_Zoom.BTSql
Imports BT_Zoom.Delegates
Imports BT_Zoom.Enums.BTSql
Imports BT_Zoom.Interfaces
Imports BT_Zoom.Constants
Imports BT_Zoom.Enums.BTDebug

'NOTE: I added this class and the shared methods as pass throughs from the methods on BTBaseList because
'          New Relic doesn't allow instrumenting of methods on a generic class.
Public Class BTBaseListHelper

    Public Shared Sub LoadDataTable(Of TTable As {ISqlBaseTableExpression, New})(list As BTBaseList(Of TTable), sql As String)

        GetData(list, sql)
        ApplyTransforms(list)
        RemoveDependentColumns(list)
        EnsureCorrectColumnOrder(list)

    End Sub

    Private Shared Sub GetData(Of TTable As {ISqlBaseTableExpression, New})(list As BTBaseList(Of TTable), sql As String)

        Dim dt As DataTable
        If list.LogMaxRowsThreshold.HasValue AndAlso list.LogMaxRowsThreshold.Value > 0 Then
            dt = DataAccessHandler.GetDataTable(sql, CommandType.Text, "tbl", list.LogMaxRowsThreshold.Value, list.Parameters.ToArray())
        Else
            dt = DataAccessHandler.GetDataTable(sql, CommandType.Text, "tbl", list.Parameters.ToArray())
        End If

        list.SetData(dt)

        If list.MaxRowsThreshold.HasValue AndAlso list.Data.Rows.Count > list.MaxRowsThreshold.Value Then
            Throw New BTException(String.Format("{0} total items were returned and the maximum allowed is {1}. Please refine your search filters.", list.Data.Rows.Count, list.MaxRowsThreshold.Value), BTException.BTExceptionType.TooManyRows)
        End If

    End Sub

    Private Shared Sub ApplyTransforms(Of TTable As {ISqlBaseTableExpression, New})(list As BTBaseList(Of TTable))
        Dim transforms As List(Of IBTSqlTransformBase) = list.Transforms
        If transforms.Count > 0 Then 'don't get this if no transforms. do not change this, as it was put in place to prevent a stack overflow (bt-15982)

            Dim sw As Stopwatch = Nothing

            For Each transform As IBTSqlTransformBase In transforms
                ApplyTransform(list, transform)
            Next

        End If

    End Sub

    <System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Reliability", "CA2000:Dispose objects before losing scope")>
    Private Shared Sub ApplyTransform(Of TTable As {ISqlBaseTableExpression, New})(list As BTBaseList(Of TTable), transform As IBTSqlTransformBase)
        Dim col As New DataColumn(transform.Alias, transform.Type)
        list.Data.Columns.Add(col)

        For i As Integer = 0 To list.Data.Rows.Count - 1
            Dim dr As DataRow = list.Data.Rows(i)
            Dim v As String = transform.ApplyTransform(dr, list)
            If v Is Nothing Then
                dr(col.Ordinal) = DBNull.Value
            ElseIf transform.Type Is GetType(String) Then
                dr(col.Ordinal) = v
            Else
                dr(col.Ordinal) = CTypeDynamic(v, transform.Type)
            End If
        Next

    End Sub

    Private Shared Sub RemoveDependentColumns(Of TTable As {ISqlBaseTableExpression, New})(list As BTBaseList(Of TTable))

        For Each expression As ISqlSelectExpression In list.SqlBuilder.SelectDependentList
            If list.SqlBuilder.SelectListAppliedToPageQuery.FindIndex(Function(x As ISqlSelectExpression) x.GetDataRowColumnName() = expression.GetDataRowColumnName()) < 0 Then
                list.Data.Columns.RemoveAt(BTSqlUtility.GetColumnIndexFromSelectExpression(list.Data, expression))
            End If
        Next

    End Sub

    Private Shared Sub EnsureCorrectColumnOrder(Of TTable As {ISqlBaseTableExpression, New})(list As BTBaseList(Of TTable))

        For i As Integer = 0 To list.OrderedSelectList.Count - 1
            Dim expr As ISqlSelectExpression = list.OrderedSelectList(i)
            Dim colName As String = expr.GetDataRowColumnName()
            list.Data.Columns(colName).SetOrdinal(i)
        Next

    End Sub

End Class

<Serializable()>
Public MustInherit Class BTBaseList(Of TTable As {ISqlBaseTableExpression, New})
    Implements IBTList, IBTBaseList(Of TTable)

    Public Sub New(ByVal fromTable As TTable)
        _sqlBuilder = New BTSqlSelectBuilder(Of TTable)(fromTable)
    End Sub

    Public Sub New()
        _sqlBuilder = New BTSqlSelectBuilder(Of TTable)(CreateTable())
    End Sub

    Protected MustOverride Function CreateTable() As TTable

    ''' <summary>
    ''' Throws an exception
    ''' </summary>
    Public Property MaxRowsThreshold As Integer? Implements IBTList.MaxRowsThreshold

    ''' <summary>
    ''' Log it
    ''' </summary> 
    Public Property LogMaxRowsThreshold As Integer? Implements IBTList.LogMaxRowsThreshold

    Private _data As DataTable
    Public ReadOnly Property Data As DataTable Implements IBTList.Data
        Get
            Return _data
        End Get
    End Property

    Friend Sub SetData(dt As DataTable)
        _data = dt
    End Sub

    Private _relatedData As New Dictionary(Of String, IBTList)
    Public ReadOnly Property RelatedData As Dictionary(Of String, IBTList) Implements IBTList.RelatedData
        Get
            Return _relatedData
        End Get
    End Property

    Private _totalRowsAllPages As Integer
    Public ReadOnly Property TotalRowsAllPages As Integer Implements IBTList.TotalRowsAllPages
        Get
            Return _totalRowsAllPages
        End Get
    End Property

    Private _lastRowReturned As Integer
    Public ReadOnly Property LastRowReturned As Integer Implements IBTList.LastRowReturned
        Get
            Return _lastRowReturned
        End Get
    End Property

    Private _firstRowReturned As Integer
    Public ReadOnly Property FirstRowReturned As Integer Implements IBTList.FirstRowReturned
        Get
            Return _firstRowReturned
        End Get
    End Property

    Private _infiniteScrollStatus As InfiniteScrollStatusTypes
    ''' <summary>
    ''' Used when using the <see cref="LoadInfiniteScroll"></see>
    ''' </summary>
    Public ReadOnly Property InfiniteScrollStatus As InfiniteScrollStatusTypes Implements IBTList.InfiniteScrollStatus
        Get
            Return _infiniteScrollStatus
        End Get
    End Property

    ''' <summary>
    ''' Used when using the <see cref="LoadInfiniteScroll"></see>
    ''' </summary>
    Public ReadOnly Property InfiniteScrollIsLoaded As Boolean Implements IBTList.InfiniteScrollIsLoaded
        Get
            If Data Is Nothing Then
                Return False
            End If

            Dim totalRows As Integer = 0
            Dim lastRow As Integer = 0

            If Data.Rows.Count > 0 Then
                totalRows = _totalRowsAllPages
                lastRow = _lastRowReturned
            End If

            Return (totalRows = lastRow)
        End Get
    End Property

    Public Property IsDistinct As Boolean Implements IBTList.IsDistinct
        Get
            Return _sqlBuilder.IsDistinct
        End Get
        Set(value As Boolean)
            _sqlBuilder.IsDistinct = value
        End Set
    End Property

    Public Property PagingSortDirection As DirectionTypes Implements IBTList.PagingSortDirection
        Get
            Return _sqlBuilder.PagingSortDirection
        End Get
        Set(value As DirectionTypes)
            _sqlBuilder.PagingSortDirection = value
        End Set
    End Property

    ''' <summary>
    ''' Get the base builder select list
    ''' </summary>
    ''' <returns>List of ISqlSelectExpressions</returns>
    Protected ReadOnly Property SelectList As List(Of ISqlSelectExpression) Implements IBTList.SelectList
        Get
            Return _sqlBuilder.SelectList
        End Get
    End Property

    Public ReadOnly Property ColumnAliasMappings As List(Of IBTSqlColumnBase) Implements IBTList.ColumnAliasMappings
        Get
            Dim result As New List(Of IBTSqlColumnBase)()
            For Each expr As ISqlSelectExpression In SelectList
                Dim col As IBTSqlColumnBase = TryCast(expr, IBTSqlColumnBase)
                If col IsNot Nothing Then
                    If Not String.IsNullOrWhitespace(col.Alias) AndAlso col.Alias <> col.Name Then
                        result.Add(col)
                    End If
                End If
            Next
            Return result
        End Get
    End Property

    ''' <summary>
    ''' Get a SelectList Dictionary
    ''' </summary>
    ''' <returns>Dictionary (of String, ISqlSelectExpression). String values correspond to the ISqlSelectExpression aliases</returns>
    ''' <remarks></remarks>
    Public ReadOnly Property SelectDictionary As Dictionary(Of String, ISqlSelectExpression) Implements IBTList.SelectDictionary
        Get
            ''Return a dict of our selectList with keys corresponding to the dataRowColumnname()
            Return Me.SelectList.ToDictionary(Function(x) x.GetDataRowColumnName())
        End Get
    End Property

#Region "SqlParameter"

    <NonSerialized>
    Private _parameters As New Dictionary(Of String, IBTSqlParameter)
    Public ReadOnly Property Parameters As IEnumerable(Of SqlParameter) Implements IBTList.Parameters
        Get
            Dim params As SqlParameter() = _parameters.Values.Select(Function(x As IBTSqlParameter) x.Parameter).ToArray()
            Dim sb As New StringBuilder()
            sb.AppendLine("--------------------------")
            DataAccessHandler.WriteOutParameters(sb, params, vbCrLf)
            sb.AppendLine("--------------------------")
            BTDebug.WriteLine(BTDebugOutputTypes.Sql, sb.ToString())
            Return params
        End Get
    End Property

    Public Function AddParameter(ByVal param As IBTSqlParameter) As IBTSqlParameter Implements IBTList.AddParameter
        If Not ContainsParameter(param.Parameter.ParameterName) Then 'to make sure we don't get parameters duplicated
            _parameters.Add(param.Parameter.ParameterName, param)
        Else
            Dim p As SqlParameter = _parameters(param.Parameter.ParameterName).Parameter
            If (p.Value Is Nothing AndAlso param.Parameter.Value Is Nothing) OrElse p.Value.Equals(param.Parameter.Value) Then
                Return _parameters(param.Parameter.ParameterName)
            End If
            Throw New Exception(String.Format("Attempting to add a duplicate sql parameter with a different value. Name: {0}", param.Parameter.ParameterName))
        End If
        Return param
    End Function

    Public Function AddParameter(ByVal paramName As String, ByVal paramValue As Integer) As IBTSqlParameter Implements IBTList.AddParameter
        Return AddParameter(DataAccessHandler.CreateSqlParameter(paramName, paramValue).ToBTSqlParameter())
    End Function

    Public Function AddParameter(ByVal paramName As String, ByVal paramValue As Integer?) As IBTSqlParameter Implements IBTList.AddParameter
        Return AddParameter(DataAccessHandler.CreateSqlParameter(paramName, paramValue).ToBTSqlParameter())
    End Function

    Public Function AddParameter(ByVal paramName As String, ByVal paramValue As SqlInt32) As IBTSqlParameter Implements IBTList.AddParameter
        Return AddParameter(DataAccessHandler.CreateSqlParameter(paramName, paramValue).ToBTSqlParameter())
    End Function

    Public Function AddParameter(ByVal paramName As String, ByVal paramValue As Boolean) As IBTSqlParameter Implements IBTList.AddParameter
        Return AddParameter(DataAccessHandler.CreateSqlParameter(paramName, paramValue).ToBTSqlParameter())
    End Function

    Public Function AddParameter(ByVal paramName As String, ByVal paramValue As Long) As IBTSqlParameter Implements IBTList.AddParameter
        Return AddParameter(DataAccessHandler.CreateSqlParameter(paramName, paramValue).ToBTSqlParameter())
    End Function

    Public Function AddParameter(ByVal paramName As String, ByVal paramValue As IEnumerable(Of Integer)) As IBTSqlParameter Implements IBTList.AddParameter
        Return AddParameter(DataAccessHandler.CreateSqlParameter(paramName, paramValue).ToBTSqlParameter())
    End Function

    Public Function AddParameter(ByVal paramName As String, ByVal paramValue As IEnumerable(Of String)) As IBTSqlParameter Implements IBTList.AddParameter
        Return AddParameter(DataAccessHandler.CreateSqlParameter(paramName, paramValue).ToBTSqlParameter())
    End Function

    Public Function AddParameter(ByVal paramName As String, ByVal paramValue As IEnumerable(Of Long)) As IBTSqlParameter Implements IBTList.AddParameter
        Return AddParameter(DataAccessHandler.CreateSqlParameter(paramName, paramValue).ToBTSqlParameter())
    End Function

    Public Function AddParameter(ByVal paramName As String, ByVal paramValue As DateTime) As IBTSqlParameter Implements IBTList.AddParameter
        Return AddParameter(DataAccessHandler.CreateSqlParameter(paramName, paramValue).ToBTSqlParameter())
    End Function

    Public Function AddParameter(ByVal paramName As String, ByVal paramValue As SqlDateTime) As IBTSqlParameter Implements IBTList.AddParameter
        Return AddParameter(DataAccessHandler.CreateSqlParameter(paramName, paramValue).ToBTSqlParameter())
    End Function

    Public Function AddParameter(ByVal paramName As String, ByVal paramValue As String, Optional varcharLength As Integer = 0) As IBTSqlParameter Implements IBTList.AddParameter
        Return AddParameter(DataAccessHandler.CreateSqlParameter(paramName, paramValue, varcharLength).ToBTSqlParameter())
    End Function

    Public Function AddParameter(ByVal paramName As String, ByVal paramValue As Decimal) As IBTSqlParameter Implements IBTList.AddParameter
        Return AddParameter(DataAccessHandler.CreateSqlParameter(paramName, paramValue).ToBTSqlParameter())
    End Function

    Public Sub AddParameters(ByVal lstParams As List(Of SqlParameter)) Implements IBTList.AddParameters
        For Each param As SqlParameter In lstParams
            AddParameter(param.ToBTSqlParameter())
        Next
    End Sub

    Public Sub AddParameters(ByVal lstParams As List(Of IBTSqlParameter)) Implements IBTList.AddParameters
        For Each param As IBTSqlParameter In lstParams
            AddParameter(param)
        Next
    End Sub

    Public Function ContainsParameter(ByVal paramName As String) As Boolean Implements IBTList.ContainsParameter
        Return _parameters.ContainsKey(paramName)
    End Function

#End Region

#Region "Load Methods -- all return the count"

    Public Function LoadAll() As Integer Implements IBTList.LoadAll
        _sql = BuildSql(False, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing)
        LoadDataTable(_sql)
        _totalRowsAllPages = Data.Rows.Count
        _firstRowReturned = 1
        _lastRowReturned = Data.Rows.Count
        Return _totalRowsAllPages
    End Function

    Public Function LoadTop(ByVal topNumToLoad As Integer) As Integer Implements IBTList.LoadTop
        _sql = BuildSql(False, topNumToLoad, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing)
        LoadDataTable(_sql)
        _totalRowsAllPages = Data.Rows.Count
        _firstRowReturned = 1
        _lastRowReturned = Data.Rows.Count
        Return _totalRowsAllPages
    End Function

    Public Function LoadPage(ByVal pageNumber As Integer, ByVal pageSize As Integer) As Integer Implements IBTList.LoadPage
        _sql = BuildSql(False, Nothing, pageNumber, pageSize, Nothing, Nothing, Nothing, Nothing)
        LoadDataTable(_sql)

        Dim returnedCount As Integer = Data.Rows.Count
        If returnedCount > 0 Then
            _totalRowsAllPages = _data.Rows(0).BT_CInt(BTConstants.DAL.PagingTotalRowsColumnName)
            'If loading data in the "up" direction (i.e. past items in the Calendar Agenda service)
            'the last row returned is the first row in the data table instead of the last row
            'because of how the list gets sorted by rowId in descending order
            If PagingSortDirection = DirectionTypes.ASC Then
                _firstRowReturned = _data.Rows.Cast(Of DataRow)().First().BT_CInt(BTConstants.DAL.PagingRowIdColumnName)
                _lastRowReturned = _data.Rows.Cast(Of DataRow)().Last().BT_CInt(BTConstants.DAL.PagingRowIdColumnName)
            ElseIf PagingSortDirection = DirectionTypes.DESC Then
                _firstRowReturned = _data.Rows.Cast(Of DataRow)().Last().BT_CInt(BTConstants.DAL.PagingRowIdColumnName)
                _lastRowReturned = _data.Rows.Cast(Of DataRow)().First().BT_CInt(BTConstants.DAL.PagingRowIdColumnName)
            End If
            'TODO: do we want to remove that column from the dataset?
        End If
        Return returnedCount
    End Function

    Public Function LoadRows(ByVal rowStartNum As Integer, ByVal rowEndNum As Integer) As Integer Implements IBTList.LoadRows
        _sql = BuildSql(False, Nothing, Nothing, Nothing, rowStartNum, rowEndNum, Nothing, Nothing)
        LoadDataTable(_sql)

        Dim returnedCount As Integer = Data.Rows.Count
        If returnedCount > 0 Then
            _totalRowsAllPages = _data.Rows(0).BT_CInt(BTConstants.DAL.PagingTotalRowsColumnName)
            'If loading data in the "up" direction (i.e. past items in the Calendar Agenda service)
            'the last row returned is the first row in the data table instead of the last row
            'because of how the list gets sorted by rowId in descending order
            If PagingSortDirection = DirectionTypes.ASC Then
                _firstRowReturned = _data.Rows.Cast(Of DataRow)().First().BT_CInt(BTConstants.DAL.PagingRowIdColumnName)
                _lastRowReturned = _data.Rows.Cast(Of DataRow)().Last().BT_CInt(BTConstants.DAL.PagingRowIdColumnName)
            ElseIf PagingSortDirection = DirectionTypes.DESC Then
                _firstRowReturned = _data.Rows.Cast(Of DataRow)().Last().BT_CInt(BTConstants.DAL.PagingRowIdColumnName)
                _lastRowReturned = _data.Rows.Cast(Of DataRow)().First().BT_CInt(BTConstants.DAL.PagingRowIdColumnName)
            End If
            'TODO: do we want to remove that column from the dataset?
        End If
        Return returnedCount
    End Function

    ''' <summary>
    ''' return first the ROW_NUMBER of the record that matches the entity ID, the _data property will contain all rows that match
    ''' </summary>
    Public Function LoadRowNum(ByVal entityId As Integer, ByVal entityIdColumn As ISqlSelectExpression) As Integer? Implements IBTList.LoadRowNum
        _sql = BuildSql(False, Nothing, Nothing, Nothing, Nothing, Nothing, entityId, entityIdColumn)
        LoadDataTable(_sql)
        Dim returnedCount As Integer = Data.Rows.Count
        If returnedCount > 0 Then
            _totalRowsAllPages = _data.Rows(0).BT_CInt(BTConstants.DAL.PagingTotalRowsColumnName)
            _firstRowReturned = _data.Rows.Cast(Of DataRow)().First().BT_CInt(BTConstants.DAL.PagingRowIdColumnName)
            _lastRowReturned = _data.Rows.Cast(Of DataRow)().Last().BT_CInt(BTConstants.DAL.PagingRowIdColumnName)
            Return _data.Rows(0).BT_CInt(BTConstants.DAL.PagingRowIdColumnName)
        End If
        Return Nothing
    End Function

    ''' <summary>
    ''' <see cref="Data"></see> contains the next page as long as the lastDisplayedEntityId matches, otherwise returns the first page and sets the reset flag
    ''' </summary>
    ''' <param name="beginRowNum">Row_number of first row to be returned</param>
    ''' <param name="endRowNum">Row_number of last row to be returned</param>
    ''' <param name="lastDisplayedEntityId">The id of the last item in the previous page</param>
    ''' <param name="entityIdColumn">The select expression that should be used to look up the entity id to match</param>
    ''' <returns>Row count returned in Data</returns>
    Public Function LoadInfiniteScroll(ByVal beginRowNum As Integer, ByVal endRowNum As Integer, ByVal lastDisplayedEntityId As Integer, ByVal entityIdColumn As ISqlSelectExpression) As Integer Implements IBTList.LoadInfiniteScroll
        _infiniteScrollStatus = InfiniteScrollStatusTypes.DataMatch
        ' 1. Load from (beginRowNum - 1) to endRowNum.  Only subtract one from beginRowNum if it is greater than one.
        Dim returnedCount As Integer = LoadRows(If(beginRowNum = 1, beginRowNum, beginRowNum - 1), endRowNum)
        ' 2a. If it's the first bunch, just return the count
        If beginRowNum = 1 Then
            Return returnedCount
        End If
        ' 2b. Compare the value of entityIdColumn in the first row of the DataTable against lastDisplayedEntityId
        If returnedCount > 0 Then
            'If loading data in the "up" direction (i.e. past items in the Calendar Agenda service)
            'the previousEntity is the last row in the data table instead of the first row
            Dim previousEntityIndex As Integer = If(PagingSortDirection = DirectionTypes.ASC, 0, returnedCount - 1)
            Dim firstEntityId As Integer = _data.Rows(previousEntityIndex).BT_CInt(entityIdColumn.GetDataRowColumnName)
            ' 3. If it matches, remove the row from the DataTable and then return the count
            If firstEntityId = lastDisplayedEntityId Then
                _data.Rows.RemoveAt(previousEntityIndex)
                Return _data.Rows.Count
            End If
        End If
        ' 4. LoadRowNum to get the foundRowNum of the record that matches lastDisplayedEntityId
        Dim foundRowNum As Integer? = LoadRowNum(lastDisplayedEntityId, entityIdColumn)
        ' 5. If it isn't found, return the first page and update the reset flag
        If Not foundRowNum.HasValue Then
            _infiniteScrollStatus = InfiniteScrollStatusTypes.ResetToStart
            Return LoadPage(1, endRowNum - beginRowNum + 1)
        End If
        ' 6. Otherwise, re-run the original query from (foundRowNum + 1) to (endRowNum - beginRowNum) + (foundRowNum + 1) and return the count
        _infiniteScrollStatus = InfiniteScrollStatusTypes.DataChanged
        Dim newBeginRowNum As Integer = foundRowNum.Value + 1
        Return LoadRows(newBeginRowNum, (endRowNum - beginRowNum) + newBeginRowNum)
    End Function

    Public Function LoadCount() As Integer Implements IBTList.LoadCount
        _sql = BuildSql(True, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing)
        _totalRowsAllPages = CInt(DataAccessHandler.ExecuteScalar(_sql, Parameters.ToArray()))
        Return _totalRowsAllPages
    End Function

    Private Sub LoadDataTable(sql As String)

        If _data IsNot Nothing Then
            _data = Nothing  'Do we want to dispose here?
        End If

        BTBaseListHelper.LoadDataTable(Me, sql)

    End Sub

    ''' <summary>
    ''' Can Only be called after Infinitescroll has been called
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub RemoveInfiniteScrollColumns() Implements IBTList.RemoveInfiniteScrollColumns
        DALManager.RemoveInfiniteScrollColumns(Data)
    End Sub

#End Region

#Region "SQL"

    'TODO: consider how to allow for CTE preceding the actual sql statement

    Private _sql As String = String.Empty
    Public ReadOnly Property Sql As String Implements IBTList.Sql
        Get
            Return _sql
        End Get
    End Property

    Public ReadOnly Property SqlForWeb As String Implements IBTList.SqlForWeb
        Get
            Return Sql.Replace(Environment.NewLine, Environment.NewLine + "<br />")
        End Get
    End Property

    Public ReadOnly Property DependencyIdentifiers As List(Of String) Implements IBTList.DependencyIdentifiers
        Get
            Return _sqlBuilder.GetDependencyIdentifiers()
        End Get
    End Property

    <NonSerialized()>
    Private _sqlBuilder As IBTSqlSelectBuilder(Of TTable)
    Public ReadOnly Property SqlBuilder As IBTSqlSelectBuilder(Of TTable) Implements IBTBaseList(Of TTable).SqlBuilder
        Get
            Return _sqlBuilder
        End Get
    End Property

    Protected ReadOnly Property SqlBuilderColumnsForMappingToEntity As List(Of IBTSqlColumnBase)
        Get
            Dim result As New List(Of IBTSqlColumnBase)()
            If _sqlBuilder IsNot Nothing Then
                For Each se As ISqlSelectExpression In SelectList
                    Dim c As IBTSqlColumnBase = TryCast(se, IBTSqlColumnBase)
                    If c IsNot Nothing Then
                        If Not String.IsNullOrWhitespace(c.Alias) AndAlso c.Alias <> c.Name Then
                            result.Add(c)
                        End If
                    End If
                Next
            End If
            Return result
        End Get
    End Property

    Private Function BuildSql(ByVal isCount As Boolean, ByVal topNumToLoad As Integer?, ByVal pageNumber As Integer?, ByVal pageSize As Integer?, ByVal rowStartNum As Integer?, ByVal rowEndNum As Integer?, ByVal entityId As Integer?, ByVal entityIdExpression As ISqlSelectExpression) As String Implements IBTList.BuildSql
        For Each transform As BTSqlTransformBase In _transforms
            _sqlBuilder.AddTransform(transform) 'will add expressions to dependent select list if they're not already in the select list or the dependent select list
        Next
        Dim dependencyIdentifiers As List(Of String) = _sqlBuilder.GetDependencyIdentifiers()
        _sqlBuilder.IsCount = isCount
        _sqlBuilder.TopNum = topNumToLoad
        _sqlBuilder.EntityId = entityId
        _sqlBuilder.EntityIdExpression = entityIdExpression

        If pageNumber.HasValue AndAlso pageSize.HasValue Then
            _sqlBuilder.SetPagingByNumberAndSize(pageNumber, pageSize)
        Else
            _sqlBuilder.SetPagingByRows(rowStartNum, rowEndNum)
        End If

        _sqlBuilder.ParametersNeededByBuilder.Clear()
        Dim s As String = _sqlBuilder.Render()
        For Each p As IBTSqlParameter In _sqlBuilder.ParametersNeededByBuilder
            If ContainsParameter(p.Parameter.ParameterName) Then
                _parameters.Remove(p.Parameter.ParameterName)
            End If
        Next
        AddParameters(_sqlBuilder.ParametersNeededByBuilder)

        Return s
    End Function

#Region "Select / Transform"

    Private _orderedSelectList As New List(Of ISqlSelectExpression)
    Public ReadOnly Property OrderedSelectList As List(Of ISqlSelectExpression) Implements IBTList.OrderedSelectList
        Get
            Return _orderedSelectList
        End Get
    End Property

    Public Sub AddSelect(ByVal ParamArray expressions() As ISqlSelectExpression) Implements IBTList.AddSelect
        If expressions IsNot Nothing AndAlso expressions.Count > 0 Then
            _sqlBuilder.AddSelect(expressions)
            _orderedSelectList.AddRange(expressions)
        End If
    End Sub

    Public Sub AddSelect(ByVal fnType As FunctionTypes, ByVal [alias] As String, ByVal ParamArray dependencies() As ISqlSelectExpression) Implements IBTList.AddSelect
        AddSelect(New BTSqlFunctionExpression(fnType, [alias], dependencies))
    End Sub

    Public Sub AddSelect(ByVal fnType As FunctionTypes, ByVal ParamArray dependencies() As ISqlSelectExpression) Implements IBTList.AddSelect
        AddSelect(New BTSqlFunctionExpression(fnType, dependencies))
    End Sub

    Public Sub AddSelect(ByVal fn As IBTSqlFunction, ByVal [alias] As String, ByVal ParamArray dependencies() As ISqlSelectExpression) Implements IBTList.AddSelect
        AddSelect(New BTSqlFunctionExpression(fn, [alias], dependencies))
    End Sub

    Public Sub AddSelect(ByVal fn As IBTSqlFunction, ByVal ParamArray dependencies() As ISqlSelectExpression) Implements IBTList.AddSelect
        AddSelect(New BTSqlFunctionExpression(fn, dependencies))
    End Sub

    Private _transforms As New List(Of IBTSqlTransformBase)
    Public ReadOnly Property Transforms As List(Of IBTSqlTransformBase) Implements IBTList.Transforms
        Get
            Return _transforms
        End Get
    End Property
    Public Sub AddTransform(ByVal ParamArray transforms() As IBTSqlTransformBase) Implements IBTList.AddTransform
        If transforms IsNot Nothing AndAlso transforms.Count > 0 Then
            _transforms.AddRange(transforms.ToArray())
            _orderedSelectList.AddRange(transforms)
        End If
    End Sub

    Public Sub AddTransform(ByVal transform As ApplySqlTransformationDelegate,
                                ByVal [alias] As String,
                                ByVal expressionList As List(Of ISqlSelectExpression),
                                ByVal sortExpressions As IBTSqlOrderByGroup,
                                Optional ByVal type As Type = Nothing) Implements IBTList.AddTransform
        If type Is Nothing Then
            type = GetType(String)
        End If

        AddTransform(New BTSqlNewColumnTransform(transform, [alias], expressionList, type, sortExpressions))
    End Sub

    ''' <summary>
    ''' Add a transform with the default sort of the first column from <paramref name="expressionList"></paramref>
    ''' </summary>
    Public Sub AddTransform(ByVal transform As ApplySqlTransformationDelegate,
                                ByVal [alias] As String,
                                ByVal expressionList As List(Of ISqlSelectExpression),
                                Optional ByVal type As Type = Nothing) Implements IBTList.AddTransform

        AddTransform(transform, [alias], expressionList, New BTSqlOrderByGroup(expressionList(0)), type)
    End Sub

    Public Sub AddTransform(ByVal transform As ApplySqlTransformationRelatedDataDelegate,
                            ByVal loadData As LoadRelatedDataForTransform,
                            ByVal [alias] As String,
                            ByVal expressionList As List(Of ISqlSelectExpression),
                            ByVal sortExpressions As IBTSqlOrderByGroup) Implements IBTList.AddTransform
        AddTransform(New BTSqlNewRelatedDataColumnTransform(transform, loadData, [alias], expressionList, sortExpressions))
    End Sub

    ''' <summary>
    ''' Add a transform with the default sort of the first column from <paramref name="expressionList"></paramref>
    ''' </summary>
    Public Sub AddTransform(ByVal transform As ApplySqlTransformationRelatedDataDelegate,
                            ByVal loadData As LoadRelatedDataForTransform,
                            ByVal [alias] As String,
                            ByVal expressionList As List(Of ISqlSelectExpression)) Implements IBTList.AddTransform
        AddTransform(New BTSqlNewRelatedDataColumnTransform(transform, loadData, [alias], expressionList))
    End Sub

#End Region

#Region "From"

    Public ReadOnly Property From As TTable Implements IBTBaseList(Of TTable).From
        Get
            Return _sqlBuilder.From
        End Get
    End Property

#End Region

#Region "Join"

    Public Function AddJoin(ByVal join As ISqlJoinExpression) As ISqlJoinExpression Implements IBTList.AddJoin
        _sqlBuilder.AddJoin(join)
        Return join
    End Function

    ''' <summary>
    ''' Creates a new join expression and adds it to the sql builder to be rendered
    ''' </summary>
    ''' <typeparam name="TJoinTable">CTE or Table type of the table parameter</typeparam>
    ''' <param name="joinType">Type of the join (e.g. Inner, Right, Left)</param>
    ''' <param name="table">Table to join with</param>
    ''' <param name="column">Column from <paramref name="table"></paramref> to be used in the ON clause</param>
    ''' <param name="otherColumn">Column from the other table we are joining with</param>
    ''' <param name="filters">Additional filters used in the ON clause</param>
    ''' <returns></returns>
    ''' <remarks>The column parameter should be a column from the TJoinTable used as the table parameter.  The other column can be a column from any other table
    ''' used in the query (the base table of the list, or another previously joined table).  Failure to put the columns in this order can cause issues
    ''' with alias renaming when using paging and infinite scroll options.</remarks>
    Public Function AddJoin(Of TJoinTable As {ISqlBaseTableExpression})(ByVal joinType As JoinTypes, ByVal table As TJoinTable, ByVal column As IBTSqlColumnBase, ByVal otherColumn As ISqlSelectExpression, ByVal ParamArray filters() As ISqlFilterExpression) As ISqlJoinExpression Implements IBTList.AddJoin
        Return AddJoin(New BTSqlTableJoin(Of TJoinTable)(joinType, table, column, otherColumn, filters))
    End Function

    Public Sub AddJoins(ByVal ParamArray joinList As ISqlJoinExpression()) Implements IBTList.AddJoins
        For Each j As ISqlJoinExpression In joinList
            AddJoin(j)
        Next
    End Sub

    Public Function AddJoinOuterApply(ByVal tbl As IBTSqlGenericFromTable) As ISqlJoinExpression Implements IBTList.AddJoinOuterApply
        Return _sqlBuilder.AddJoinOuterApply(tbl)
    End Function

#End Region

#Region "Filter"

    Public Property AllowNoFilters As Boolean Implements IBTList.AllowNoFilters
        Get
            Return _sqlBuilder.AllowNoFilters
        End Get
        Set(value As Boolean)
            _sqlBuilder.AllowNoFilters = value
        End Set
    End Property

    Public Property EnableOptionRecompile() As Boolean Implements IBTList.EnableOptionRecompile
        Get
            Return _sqlBuilder.EnableOptionRecompile
        End Get
        Set(value As Boolean)
            _sqlBuilder.EnableOptionRecompile = value
        End Set
    End Property

    Public Property EnableOptionOptimizeForUnknown() As Boolean Implements IBTList.EnableOptionOptimizeForUnknown
        Get
            Return _sqlBuilder.EnableOptionOptimizeForUnknown
        End Get
        Set(value As Boolean)
            _sqlBuilder.EnableOptionOptimizeForUnknown = value
        End Set
    End Property

    Public Property EnableOptionForceOrder() As Boolean Implements IBTList.EnableOptionForceOrder
        Get
            Return _sqlBuilder.EnableOptionForceOrder
        End Get
        Set(value As Boolean)
            _sqlBuilder.EnableOptionForceOrder = value
        End Set
    End Property

    Public Sub AddFilter(ByVal filter As ISqlFilterExpression) Implements IBTList.AddFilter
        _sqlBuilder.AddFilter(filter)
    End Sub

    Public Sub AddFilters(ParamArray filterList As ISqlFilterExpression()) Implements IBTList.AddFilters
        For Each f As ISqlFilterExpression In filterList
            AddFilter(f)
        Next
    End Sub

    Public Function AddFilterGroup(ByVal booleanOperator As BooleanOperatorTypes, ParamArray filterList As ISqlFilterExpression()) As IBTSqlFilterGroup Implements IBTList.AddFilterGroup
        Return _sqlBuilder.AddFilterGroup(booleanOperator, filterList)
    End Function

    Public Sub AddFilter(ByVal leftExpression As ISqlSelectExpression, ByVal comparisonOperator As ComparisonOperatorTypes, ByVal rightClause As String, Optional ByVal booleanOperator As BooleanOperatorTypes = BooleanOperatorTypes.AND) Implements IBTList.AddFilter
        AddFilter(New BTSqlComparisonFilter(leftExpression, comparisonOperator, rightClause, booleanOperator))
    End Sub
    Public Sub AddFilter(ByVal leftExpression As ISqlSelectExpression, ByVal comparisonOperator As ComparisonOperatorTypes, ByVal rightExpression As ISqlSelectExpression, Optional ByVal booleanOperator As BooleanOperatorTypes = BooleanOperatorTypes.AND) Implements IBTList.AddFilter
        AddFilter(New BTSqlComparisonFilter(leftExpression, comparisonOperator, rightExpression, booleanOperator))
    End Sub
    Public Sub AddFilter(ByVal leftExpression As ISqlSelectExpression, ByVal comparisonOperator As ComparisonOperatorTypes, ByVal rightExpression As IBTSqlParameter, Optional ByVal booleanOperator As BooleanOperatorTypes = BooleanOperatorTypes.AND) Implements IBTList.AddFilter
        AddFilter(New BTSqlComparisonFilter(leftExpression, comparisonOperator, rightExpression, booleanOperator))
    End Sub
    Public Sub AddFilter(ByVal leftExpression As ISqlSelectExpression, ByVal logicalOperator As LogicalOperatorTypes, Optional ByVal booleanOperator As BooleanOperatorTypes = BooleanOperatorTypes.AND) Implements IBTList.AddFilter
        AddFilter(New BTSqlLogicalFilter(leftExpression, logicalOperator, booleanOperator))
    End Sub
    Public Sub AddFilter(ByVal leftExpression As ISqlSelectExpression, ByVal logicalOperator As LogicalOperatorTypes, ByVal logicalExpression As String, Optional ByVal booleanOperator As BooleanOperatorTypes = BooleanOperatorTypes.AND) Implements IBTList.AddFilter
        AddFilter(New BTSqlLogicalFilter(leftExpression, logicalOperator, logicalExpression, booleanOperator))
    End Sub
    Public Sub AddFilter(ByVal leftExpression As ISqlSelectExpression, ByVal logicalOperator As LogicalOperatorTypes, ByVal leftBetween As IBTSqlParameter, ByVal rightBetween As IBTSqlParameter, Optional ByVal booleanOperator As BooleanOperatorTypes = BooleanOperatorTypes.AND) Implements IBTList.AddFilter
        AddFilter(New BTSqlLogicalFilter(leftExpression, logicalOperator, leftBetween, rightBetween, booleanOperator))
    End Sub
    Public Sub AddFilter(ByVal leftExpression As ISqlSelectExpression, ByVal logicalOperator As LogicalOperatorTypes, ByVal rightExpression As ISqlSelectExpression, Optional ByVal booleanOperator As BooleanOperatorTypes = BooleanOperatorTypes.AND) Implements IBTList.AddFilter
        AddFilter(New BTSqlLogicalFilter(leftExpression, logicalOperator, rightExpression, booleanOperator))
    End Sub
    Public Sub AddFilter(ByVal leftExpression As ISqlSelectExpression, ByVal logicalOperator As LogicalOperatorTypes, ByVal rightExpression As IBTSqlParameter, Optional ByVal booleanOperator As BooleanOperatorTypes = BooleanOperatorTypes.AND) Implements IBTList.AddFilter
        AddFilter(New BTSqlLogicalFilter(leftExpression, logicalOperator, rightExpression, booleanOperator))
    End Sub
    Public Sub AddFilter(ByVal logicalOperator As LogicalOperatorTypes, ByVal rightExpression As ISqlSelectQuery, Optional ByVal booleanOperator As BooleanOperatorTypes = BooleanOperatorTypes.AND) Implements IBTList.AddFilter
        AddFilter(New BTSqlLogicalFilter(logicalOperator, rightExpression, booleanOperator))
    End Sub

    ''' <summary>
    ''' Adds all the filterExpressions joined by the OR operator and surrounded by open/close parentheses
    ''' </summary>
    ''' <param name="filterExpressions">The filter expressions to be joined by OR operator</param>
    ''' <param name="booleanOperator">The boolean operator used prior to the first item</param>
    Public Function AddOrFilterGroup(ByVal filterExpressions As IEnumerable(Of IBTSqlFilter), Optional ByVal booleanOperator As BooleanOperatorTypes = BooleanOperatorTypes.AND) As IBTSqlFilterGroup Implements IBTList.AddOrFilterGroup
        Dim fg As IBTSqlFilterGroup = AddFilterGroup(booleanOperator) 'the group gets the booleanOperator supplied by the optional param
        For j As Integer = 0 To filterExpressions.Count - 1
            Dim exp As IBTSqlFilter = filterExpressions(j)
            exp.BooleanOperator = BooleanOperatorTypes.OR ' all expression set to an OR
            fg.AddFilter(exp)
        Next
        Return fg
    End Function

    ''' <summary>
    ''' Adds LIKE filters for the provided searchFields
    ''' </summary>
    ''' <remarks>Moved and adapted from BTEntityLists.AddKeywordSearchFilter</remarks>
    Public Function AddKeywordSearchFilter(ByVal parameterName As String, ByVal leftExpressions As List(Of ISqlSelectExpression), ByVal rightExpressions As List(Of String), Optional ByVal booleanOperator As BooleanOperatorTypes = BooleanOperatorTypes.AND) As IBTSqlFilterGroup Implements IBTList.AddKeywordSearchFilter
        Dim group As IBTSqlFilterGroup = _sqlBuilder.AddKeywordSearchFilter(parameterName, leftExpressions, rightExpressions, booleanOperator)
        AddParameters(_sqlBuilder.Parameters.ToList)
        Return group
    End Function

    ''' <summary>
    ''' Adds LIKE filters for the provided searchField
    ''' </summary>
    ''' <remarks>Just an overload to use a single column and string</remarks>
    Public Function AddKeywordSearchFilter(ByVal parameterName As String, ByVal leftExpression As ISqlSelectExpression, ByVal rightExpression As String, Optional ByVal booleanOperator As BooleanOperatorTypes = BooleanOperatorTypes.AND) As IBTSqlFilterGroup Implements IBTList.AddKeywordSearchFilter
        Dim lExpressionGroup As New List(Of ISqlSelectExpression) From {leftExpression}
        Dim rExpressionGroup As New List(Of String) From {rightExpression}

        Return AddKeywordSearchFilter(parameterName, lExpressionGroup, rExpressionGroup, booleanOperator)
    End Function

    Public Function GenerateKeywordSearchFilterGroup(ByVal parameterName As String, ByVal leftExpressions As List(Of ISqlSelectExpression), ByVal rightExpressions As List(Of String), Optional ByVal booleanOperator As BooleanOperatorTypes = BooleanOperatorTypes.AND) As IBTSqlFilterGroup Implements IBTList.GenerateKeywordSearchFilterGroup
        Dim group As IBTSqlFilterGroup = _sqlBuilder.GenerateKeywordSearchFilterGroup(parameterName, leftExpressions, rightExpressions, booleanOperator)
        AddParameters(_sqlBuilder.Parameters.ToList)
        Return group
    End Function

    ''' <summary>
    ''' Adds full text filters for the provided searchFields
    ''' </summary>
    ''' <remarks>Moved and adapted from BTEntityLists.AddFullTextKeywordSearchFilter</remarks>
    Public Function AddFullTextKeywordSearchFilter(ByVal parameterName As String, ByVal leftExpressions As List(Of ISqlSelectExpression), ByVal rightExpressions As List(Of String), Optional ByVal booleanOperator As BooleanOperatorTypes = BooleanOperatorTypes.AND, Optional ByVal partialWordSearch As Boolean = False) As IBTSqlFilterGroup Implements IBTList.AddFullTextKeywordSearchFilter
        Dim group As IBTSqlFilterGroup = _sqlBuilder.GenerateKeywordSearchFilterGroup(parameterName, leftExpressions, rightExpressions, booleanOperator, partialWordSearch, True)
        AddParameters(_sqlBuilder.Parameters.ToList)
        Return group
    End Function

    ''' <summary>
    ''' Creates full text filters for the provided searchFields. Does not add filter group to base list query.
    ''' </summary>
    ''' <remarks></remarks>
    Public Function GenerateFullTextKeywordSearchFilterGroup(ByVal parameterName As String, ByVal leftExpressions As List(Of ISqlSelectExpression), ByVal rightExpressions As List(Of String), Optional ByVal booleanOperator As BooleanOperatorTypes = BooleanOperatorTypes.AND, Optional ByVal partialWordSearch As Boolean = False) As IBTSqlFilterGroup Implements IBTList.GenerateFullTextKeywordSearchFilterGroup
        Dim group As IBTSqlFilterGroup = _sqlBuilder.GenerateKeywordSearchFilterGroup(parameterName, leftExpressions, rightExpressions, booleanOperator, partialWordSearch, True)
        AddParameters(_sqlBuilder.Parameters.ToList)
        Return group
    End Function

    Public Sub AddJobsiteFilter(leftExpression As ISqlSelectExpression, filter As IJobsiteFilter) Implements IBTList.AddJobsiteFilter
        If filter.ShouldUseSubQuery Then
            AddFilter(leftExpression, LogicalOperatorTypes.In, filter.SubQuery)
            AddParameters(filter.SubQuery.GetParameters().Cast(Of IBTSqlParameter).ToList())
        Else
            AddFilter(leftExpression, LogicalOperatorTypes.In, AddParameter("@jobFilter", filter.JobIDs))
        End If
    End Sub

#End Region

#Region "Group By"

    Public Sub AddGroupBy(ByVal ParamArray expressions() As ISqlGroupByExpression) Implements IBTList.AddGroupBy
        If expressions IsNot Nothing AndAlso expressions.Length > 0 Then
            _sqlBuilder.AddGroupBy(expressions)
        End If
    End Sub
    Public Sub AddGroupBy(ByVal ParamArray expressions() As ISqlSelectExpression) Implements IBTList.AddGroupBy
        If expressions IsNot Nothing AndAlso expressions.Length > 0 Then
            _sqlBuilder.AddGroupBy(expressions)
        End If
    End Sub

#End Region

#Region "Having"

    Public Sub AddHaving(ByVal filter As ISqlFilterExpression) Implements IBTList.AddHaving
        _sqlBuilder.AddHaving(filter)
    End Sub

#End Region

#Region "Order By"

    Public Sub AddOrderBy(ByVal orderBy As ISqlOrderByExpression) Implements IBTList.AddOrderBy
        _sqlBuilder.AddOrderBy(orderBy)
    End Sub
    Public Sub AddOrderBys(ByVal ParamArray orderByList As ISqlOrderByExpression()) Implements IBTList.AddOrderBys
        For Each o As ISqlOrderByExpression In orderByList
            AddOrderBy(o)
        Next
    End Sub
    Public Sub AddOrderBy(ByVal text As String, ByVal direction As DirectionTypes, ByVal ParamArray dependencies() As ISqlSelectExpression) Implements IBTList.AddOrderBy
        AddOrderBy(New BTSqlTextOrderByExpression(text, direction, dependencies))
    End Sub
    Public Sub AddOrderBy(ByVal column As ISqlSelectExpression, ByVal direction As DirectionTypes) Implements IBTList.AddOrderBy
        _sqlBuilder.AddOrderBy(New BTSqlOrderBy(column, direction))
    End Sub
    Public Sub AddOrderBy(ByVal fnType As FunctionTypes, ByVal direction As DirectionTypes, ByVal ParamArray dependencies() As ISqlSelectExpression) Implements IBTList.AddOrderBy
        _sqlBuilder.AddOrderBy(fnType, direction, dependencies)
    End Sub
    Public Sub AddOrderBy(ByVal fn As IBTSqlFunction, ByVal direction As DirectionTypes, ByVal ParamArray dependencies() As ISqlSelectExpression) Implements IBTList.AddOrderBy
        _sqlBuilder.AddOrderBy(fn, direction, dependencies)
    End Sub
    Public Sub AddOrderByAscending(ByVal column As ISqlSelectExpression) Implements IBTList.AddOrderByAscending
        _sqlBuilder.AddOrderByAscending(column)
    End Sub
    Public Sub AddOrderByAscending(ByVal text As String, ByVal ParamArray dependencies() As ISqlSelectExpression) Implements IBTList.AddOrderByAscending
        AddOrderBy(New BTSqlTextOrderByExpression(text, DirectionTypes.ASC, dependencies))
    End Sub
    Public Sub AddOrderByAscending(ByVal fnType As FunctionTypes, ByVal ParamArray dependencies() As ISqlSelectExpression) Implements IBTList.AddOrderByAscending
        _sqlBuilder.AddOrderByAscending(fnType, dependencies)
    End Sub
    Public Sub AddOrderByAscending(ByVal fn As IBTSqlFunction, ByVal ParamArray dependencies() As ISqlSelectExpression) Implements IBTList.AddOrderByAscending
        _sqlBuilder.AddOrderByAscending(fn, dependencies)
    End Sub
    Public Sub AddOrderByDescending(ByVal column As ISqlSelectExpression) Implements IBTList.AddOrderByDescending
        _sqlBuilder.AddOrderByDescending(column)
    End Sub
    Public Sub AddOrderByDescending(ByVal text As String, ByVal ParamArray dependencies() As ISqlSelectExpression) Implements IBTList.AddOrderByDescending
        AddOrderBy(New BTSqlTextOrderByExpression(text, DirectionTypes.DESC, dependencies))
    End Sub
    Public Sub AddOrderByDescending(ByVal fnType As FunctionTypes, ByVal ParamArray dependencies() As ISqlSelectExpression) Implements IBTList.AddOrderByDescending
        _sqlBuilder.AddOrderByDescending(fnType, dependencies)
    End Sub
    Public Sub AddOrderByDescending(ByVal fn As IBTSqlFunction, ByVal ParamArray dependencies() As ISqlSelectExpression) Implements IBTList.AddOrderByDescending
        _sqlBuilder.AddOrderByDescending(fn, dependencies)
    End Sub

#End Region

#Region "CTE"

    Public Function AddCte(Of TOtherTable As {ISqlBaseTableExpression, New})(ByVal cte As IBTSqlCteExpression(Of TOtherTable)) As ISqlCteExpression Implements IBTList.AddCte
        Return _sqlBuilder.AddCte(cte)
    End Function


    Public Function AddCte(Of TOtherTable As {ISqlBaseTableExpression, New})(ByVal sqlBuilder As IBTSqlSelectBuilder(Of TOtherTable), ByVal cteName As String, ByVal [alias] As String, ByVal ParamArray unionQueries As ISqlUnionable()) As ISqlCteExpression Implements IBTList.AddCte
        Return _sqlBuilder.AddCte(sqlBuilder, cteName, [alias], unionQueries)
    End Function

    Public Function AddCte(Of TOtherTable As {ISqlBaseTableExpression, New})(ByVal unionBuilder As IBTSqlUnionBuilder(Of TOtherTable), ByVal cteName As String, ByVal [alias] As String) As ISqlCteExpression Implements IBTList.AddCte
        Return _sqlBuilder.AddCte(unionBuilder, cteName, [alias])
    End Function

    Public Function AddCte(Of TOtherTable As {ISqlBaseTableExpression, New})(ByVal cte As IBTSqlCteExpression(Of TOtherTable), ByVal position As Integer) As ISqlCteExpression Implements IBTList.AddCte
        Return _sqlBuilder.AddCte(cte, position)
    End Function

    Public Function AddCte(Of TOtherTable As {ISqlBaseTableExpression, New})(ByVal sqlBuilder As IBTSqlSelectBuilder(Of TOtherTable), ByVal cteName As String, ByVal [alias] As String, ByVal position As Integer, ByVal ParamArray unionQueries As ISqlUnionable()) As ISqlCteExpression Implements IBTList.AddCte
        Return _sqlBuilder.AddCte(sqlBuilder, cteName, [alias], position, unionQueries)
    End Function

    Public Function AddCte(Of TOtherTable As {ISqlBaseTableExpression, New})(ByVal unionBuilder As IBTSqlUnionBuilder(Of TOtherTable), ByVal cteName As String, ByVal [alias] As String, ByVal position As Integer) As ISqlCteExpression Implements IBTList.AddCte
        Return _sqlBuilder.AddCte(unionBuilder, cteName, [alias], position)
    End Function

#End Region

#Region "Temp tables"

    Public Function AddTempTable(ByVal tempTbl As ISqlTempTableExpression) As ISqlTempTableExpression Implements IBTList.AddTempTable
        Return _sqlBuilder.AddTempTable(tempTbl)
    End Function

    Public Function AddTempTable(Of T As {ISqlTempTableExpression, New})() As ISqlTempTableExpression Implements IBTList.AddTempTable
        Return _sqlBuilder.AddTempTable(New T())
    End Function

#End Region

#Region "AccountingIDs"

    Public Function AddAccountingIDs(fkColumn As IBTSqlColumn, tableAlias As String, columnPrefix As String) As List(Of ISqlSelectExpression) Implements IBTList.AddAccountingIDs
        Return _sqlBuilder.AddAccountingIDs(fkColumn, tableAlias, columnPrefix)
    End Function

#End Region

#End Region

#Region "IDisposable"

    Private disposedValue As Boolean ' To detect redundant calls

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                'dispose managed state (managed objects).
                If _data IsNot Nothing Then
                    _data.Dispose()
                End If

                If Not RelatedData.IsNullOrEmpty() Then
                    For Each l As IBTList In RelatedData.Values
                        If l IsNot Nothing Then 'this is probably a redundant check but we'll do it anyways, just in case
                            l.Dispose()
                        End If
                    Next
                End If
            End If

            'set large fields to null.
            _data = Nothing
            _relatedData = Nothing
            _sqlBuilder = Nothing
        End If
        Me.disposedValue = True
    End Sub

    ' This code added by Visual Basic to correctly implement the disposable pattern.
    Public Sub Dispose() Implements IDisposable.Dispose
        ' Do not change this code.  Put cleanup code in Dispose(disposing As Boolean) above.
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub

#End Region

End Class

<Serializable()>
Public MustInherit Class BTBaseEntityList(Of TEntity As {IBTBaseEntity(Of TTable, TList)}, TTable As {IBTSqlTable, New}, TList As {IBTBaseList(Of TTable)})
    Inherits BTBaseList(Of TTable)
    Implements IEnumerable(Of TEntity), IBTBaseEntityList(Of TEntity, TTable, TList)

    <NonSerialized>
    Private _entityListCache As List(Of TEntity)

    Public Property CacheEntityList As Boolean = False Implements IBTBaseEntityList(Of TEntity, TTable, TList).CacheEntityList

    Public Sub New(ByVal fromTable As TTable)
        MyBase.New(fromTable)
    End Sub

    Public Sub New()
        MyBase.New()
    End Sub

    Protected MustOverride Function CreateEntity() As TEntity

    Public Function GetEntity(ByVal idx As Integer) As TEntity Implements IBTBaseEntityList(Of TEntity, TTable, TList).GetEntity
        If Data Is Nothing Then
            Throw New NullReferenceException("GetEntity failed because Data is Nothing")
        End If
        If idx < 0 OrElse idx >= Data.Rows.Count Then
            Throw New ArgumentOutOfRangeException("idx", idx, "GetEntity failed because idx is out of range")
        End If

        'We can't just invoke the constructor that accepts a DataRow because of limitations in the way .NET Generics behave with respect to constraints.
        Dim entity As TEntity = CreateEntity()

        'We need to provide any of the select expressions that used an alias so that we can map columns in the DataRow to Fields on the Entity.
        '   e.g., BTSubsTable.Pword.WithAlias("Password") will result in a column named "Password", but the BTSubs entity has a property "Pword".
        '         PopulateFromDataRow is "smart enough" to properly map the "Password" column to the "Pword" entity property, as long as we
        '         provide the select expressions that were used to populate the DataTable and those select expressions have the aliases that were used.
        entity.SetList(Me)
        entity.PopulateFromDataRow(Data.Rows(idx), SqlBuilderColumnsForMappingToEntity.ToArray())

        Return entity
    End Function

    Public Iterator Function GetEnumerator() As IEnumerator(Of TEntity) Implements IEnumerable(Of TEntity).GetEnumerator
        If Data Is Nothing Then
            Exit Function
        End If
        If Not CacheEntityList Then
            For i As Integer = 0 To Data.Rows.Count - 1
                Yield GetEntity(i)
            Next
        ElseIf _entityListCache IsNot Nothing Then
            For Each e As TEntity In _entityListCache
                Yield e
            Next
        Else
            _entityListCache = New List(Of TEntity)
            For i As Integer = 0 To Data.Rows.Count - 1
                Dim e As TEntity = GetEntity(i)
                _entityListCache.Add(e)
                Yield e
            Next
        End If
    End Function

    Private Function GetEnumerator1() As IEnumerator Implements IEnumerable.GetEnumerator
        Return GetEnumerator()
    End Function

End Class

Public Class BTCteList(Of TTable As {ISqlBaseTableExpression, New})
    Inherits BTBaseList(Of BtSqlCteExpression(Of TTable))

    Protected Overrides Function CreateTable() As BtSqlCteExpression(Of TTable)
        Throw New NotImplementedException()
    End Function

    Private ReadOnly _cteExpression As IBTSqlCteExpression(Of TTable)

    Public Sub New(ByVal sqlCte As BtSqlCteExpression(Of TTable))
        MyBase.New(sqlCte)
        _cteExpression = sqlCte
    End Sub

    Sub New(ByVal baseSqlBuilder As BTSqlSelectBuilder(Of TTable), ByVal cteName As String, ByVal ParamArray unionQueries As ISqlUnionable())
        Me.New(baseSqlBuilder, cteName, "gcte", unionQueries)
    End Sub

    Sub New(ByVal baseSqlBuilder As BTSqlSelectBuilder(Of TTable), ByVal cteName As String, ByVal [alias] As String, ByVal ParamArray unionQueries As ISqlUnionable())
        Me.New(GetCteExpression(baseSqlBuilder, cteName, [alias], unionQueries))
    End Sub

    Private Shared Function GetCteExpression(ByVal baseSqlBuilder As BTSqlSelectBuilder(Of TTable), ByVal cteName As String, ByVal [alias] As String, ByVal ParamArray unionQueries As ISqlUnionable()) As BtSqlCteExpression(Of TTable)
        If unionQueries IsNot Nothing AndAlso unionQueries.Any Then
            'only use the union builder if we actually need it
            Return New BtSqlCteExpression(Of TTable)(New BTSqlUnionBuilder(Of TTable)(baseSqlBuilder, unionQueries), cteName, [alias])
        Else
            Return New BtSqlCteExpression(Of TTable)(baseSqlBuilder, cteName, [alias])
        End If
    End Function

    Public ReadOnly Property CteTable As IBTSqlCteExpression(Of TTable)
        Get
            Return _cteExpression
        End Get
    End Property

    Public Function CteColumn(column As ISqlSelectExpression) As IBTSqlCteColumn
        Return _cteExpression.CteColumn(column)
    End Function

    Public Sub AddUnion(ByVal sqlBuilder As ISqlSelectBuilder, ByVal unionType As UnionTypes)
        Dim q As ISqlUnionable = New BTSqlUnionSelectQuery(sqlBuilder.ToSelectQuery, unionType)
        _cteExpression.AddUnion(q)
    End Sub

End Class
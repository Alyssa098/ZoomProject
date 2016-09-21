Option Strict On
Option Explicit On

Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports System.Data.SqlClient
Imports System.Data.SqlTypes
Imports BT_Zoom.BTSql
Imports BT_Zoom.Enums.BTSql
Imports BT_Zoom.Interfaces

<Serializable>
Public MustInherit Class BTSqlBuilderBase(Of TTable As {IBaseTableExpression, New})
    Implements ISqlBuilder, IBTSqlBuilderBase(Of TTable)

    Protected Const Sql_TOP As String = "TOP"
    Protected Const Sql_FROM As String = "FROM"
    Protected Const Sql_WHERE As String = "WHERE"

    Protected Const Sql_TopMax As Integer = 500

    Protected ReadOnly _from As TTable

    Public Sub New(Optional ByVal useLineBreaks As Boolean = True)
        _from = New TTable
        Me.UseLineBreaks = useLineBreaks
    End Sub

    Public Sub New(ByVal fromTable As TTable, Optional ByVal useLineBreaks As Boolean = True)
        _from = fromTable
        Me.UseLineBreaks = useLineBreaks
    End Sub

    Public Property UseLineBreaks As Boolean = True Implements ISqlBuilder.UseLineBreaks

    Protected Sub AppendLineOrSpace(ByVal sb As StringBuilder)
        If UseLineBreaks Then
            sb.AppendLine()
        Else
            sb.Append(" ")
        End If
    End Sub

    Public Function GetParameters() As List(Of IBTSqlParameter) Implements IHasParameters.GetParameters
        Return _parameters.Values.ToList
    End Function

    Public Overridable Function Render() As String Implements ISqlExpression.Render
        Return Render(True)
    End Function

    Public Overridable Function RenderForCte() As String Implements ISqlBuilder.RenderForCte
        'if we are rendering a CTE inside a CTE, we don't want to render the WITH clause again
        Return Render(False)
    End Function

    Private Function Render(ByVal renderWithClause As Boolean) As String
        Dim sb As New StringBuilder
        sb.Append(GetSqlTempTableDeclaration())
        If renderWithClause Then
            sb.Append(GetSqlWithClause())
        End If
        Render(sb)
        sb.Append(GetSqlTempTableCleanup())
        Return sb.ToString
    End Function

    ''' <summary>
    ''' Do additional rendering after CTE
    ''' </summary>
    Public MustOverride Sub Render(ByRef sb As StringBuilder) Implements ISqlBuilder.Render

    Public Property TopNum As Integer? = Nothing Implements ISqlBuilder.TopNum

#Region "SqlParameter"

    <NonSerialized>
    Private _parameters As New Dictionary(Of String, IBTSqlParameter)
    Public ReadOnly Property Parameters As IEnumerable(Of SqlParameter) Implements IHasParameters.Parameters
        Get
            Return _parameters.Values.Select(Function(x As IBTSqlParameter) x.Parameter)
        End Get
    End Property

    Public Function AddParameter(ByVal param As IBTSqlParameter) As IBTSqlParameter Implements ISqlBuilder.AddParameter
        If Not ContainsParameter(param.Parameter.ParameterName) Then 'to make sure we don't get parameters duplicated
            _parameters.Add(param.Parameter.ParameterName, param)
        Else
            If _parameters(param.Parameter.ParameterName).Parameter.Value.Equals(param.Parameter.Value) Then
                Return DirectCast(_parameters(param.Parameter.ParameterName), BTSqlParameter)
            End If
            Throw New Exception(String.Format("Attempting to add a duplicate sql parameter with a different value. Name: {0}", param.Parameter.ParameterName))
        End If
        Return param
    End Function

    Public Function AddParameter(ByVal paramName As String, ByVal paramValue As Integer) As IBTSqlParameter Implements ISqlBuilder.AddParameter
        Return AddParameter(DataAccessHandler.CreateSqlParameter(paramName, paramValue).ToBTSqlParameter())
    End Function

    Public Function AddParameter(ByVal paramName As String, ByVal paramValue As SqlInt32) As IBTSqlParameter Implements ISqlBuilder.AddParameter
        Return AddParameter(DataAccessHandler.CreateSqlParameter(paramName, paramValue).ToBTSqlParameter())
    End Function

    Public Function AddParameter(ByVal paramName As String, ByVal paramValue As Boolean) As IBTSqlParameter Implements ISqlBuilder.AddParameter
        Return AddParameter(DataAccessHandler.CreateSqlParameter(paramName, paramValue).ToBTSqlParameter())
    End Function

    Public Function AddParameter(ByVal paramName As String, ByVal paramValue As SqlBoolean) As IBTSqlParameter Implements ISqlBuilder.AddParameter
        Return AddParameter(DataAccessHandler.CreateSqlParameter(paramName, paramValue).ToBTSqlParameter())
    End Function

    Public Function AddParameter(ByVal paramName As String, ByVal paramValue As Long) As IBTSqlParameter Implements ISqlBuilder.AddParameter
        Return AddParameter(DataAccessHandler.CreateSqlParameter(paramName, paramValue).ToBTSqlParameter())
    End Function

    Public Function AddParameter(ByVal paramName As String, ByVal paramValue As SqlInt64) As IBTSqlParameter Implements ISqlBuilder.AddParameter
        Return AddParameter(DataAccessHandler.CreateSqlParameter(paramName, paramValue).ToBTSqlParameter())
    End Function

    Public Function AddParameter(ByVal paramName As String, ByVal paramValue As IEnumerable(Of Integer)) As IBTSqlParameter Implements ISqlBuilder.AddParameter
        Return AddParameter(DataAccessHandler.CreateSqlParameter(paramName, paramValue).ToBTSqlParameter())
    End Function

    Public Function AddParameter(ByVal paramName As String, ByVal paramValue As IEnumerable(Of Long)) As IBTSqlParameter Implements ISqlBuilder.AddParameter
        Return AddParameter(DataAccessHandler.CreateSqlParameter(paramName, paramValue).ToBTSqlParameter())
    End Function

    Public Function AddParameter(ByVal paramName As String, ByVal paramValue As IEnumerable(Of String)) As IBTSqlParameter Implements ISqlBuilder.AddParameter
        Return AddParameter(DataAccessHandler.CreateSqlParameter(paramName, paramValue).ToBTSqlParameter())
    End Function

    Public Function AddParameter(ByVal paramName As String, ByVal paramValue As IEnumerable(Of DateTime)) As IBTSqlParameter Implements ISqlBuilder.AddParameter
        Return AddParameter(DataAccessHandler.CreateSqlParameter(paramName, paramValue).ToBTSqlParameter())
    End Function

    Public Function AddParameter(ByVal paramName As String, ByVal paramValue As DateTime) As IBTSqlParameter Implements ISqlBuilder.AddParameter
        Return AddParameter(DataAccessHandler.CreateSqlParameter(paramName, paramValue).ToBTSqlParameter())
    End Function

    Public Function AddParameter(ByVal paramName As String, ByVal paramValue As SqlDateTime) As IBTSqlParameter Implements ISqlBuilder.AddParameter
        Return AddParameter(DataAccessHandler.CreateSqlParameter(paramName, paramValue).ToBTSqlParameter())
    End Function

    Public Function AddParameter(ByVal paramName As String, ByVal paramValue As String, Optional varcharLength As Integer = 0) As IBTSqlParameter Implements ISqlBuilder.AddParameter
        Return AddParameter(DataAccessHandler.CreateSqlParameter(paramName, paramValue, varcharLength).ToBTSqlParameter())
    End Function

    Public Function AddParameter(ByVal paramName As String, ByVal paramValue As SqlString, Optional varcharLength As Integer = 0) As IBTSqlParameter Implements ISqlBuilder.AddParameter
        Return AddParameter(DataAccessHandler.CreateSqlParameter(paramName, paramValue, varcharLength).ToBTSqlParameter())
    End Function

    Public Function AddParameter(ByVal paramName As String, ByVal paramValue As Decimal) As IBTSqlParameter Implements ISqlBuilder.AddParameter
        Return AddParameter(DataAccessHandler.CreateSqlParameter(paramName, paramValue).ToBTSqlParameter())
    End Function

    Public Function AddParameter(ByVal paramName As String, ByVal paramValue As SqlDecimal) As IBTSqlParameter Implements ISqlBuilder.AddParameter
        Return AddParameter(DataAccessHandler.CreateSqlParameter(paramName, paramValue).ToBTSqlParameter())
    End Function

    Public Function AddParameter(ByVal paramName As String, ByVal paramValue As SqlGuid) As IBTSqlParameter Implements ISqlBuilder.AddParameter
        Return AddParameter(DataAccessHandler.CreateSqlParameter(paramName, paramValue).ToBTSqlParameter())
    End Function

    Public Sub AddParameters(ByVal lstParams As List(Of SqlParameter)) Implements ISqlBuilder.AddParameters
        For Each param As SqlParameter In lstParams
            AddParameter(param.ToBTSqlParameter())
        Next
    End Sub

    Public Sub AddParameters(ByVal lstParams As List(Of IBTSqlParameter)) Implements ISqlBuilder.AddParameters
        For Each param As IBTSqlParameter In lstParams
            AddParameter(param)
        Next
    End Sub

    Public Function ContainsParameter(ByVal paramName As String) As Boolean Implements ISqlBuilder.ContainsParameter
        Return _parameters.ContainsKey(paramName)
    End Function

#End Region

#Region "Temp Tables"

    Private ReadOnly _tempTblList As New List(Of ISqlTempTableExpression)

    Public Function AddTempTable(Of T As {ISqlTempTableExpression, New})() As ISqlTempTableExpression Implements ISqlBuilder.AddTempTable
        Return AddTempTable(New T())
    End Function

    Public Function AddTempTable(ByVal tempTbl As ISqlTempTableExpression) As ISqlTempTableExpression Implements ISqlBuilder.AddTempTable
        _tempTblList.Add(tempTbl)
        Return tempTbl
    End Function

    Private Function GetSqlTempTableDeclaration() As String
        Dim sb As New StringBuilder()

        For Each t As ISqlTempTableExpression In _tempTblList
            If t.RenderType.HasFlag(TempTableRenderTypes.Create) Then
                sb.Append(t.RenderDeclaration())
            End If
            AppendLineOrSpace(sb)
        Next

        Return sb.ToString()
    End Function

    Private Function GetSqlTempTableCleanup() As String
        Dim sb As New StringBuilder()

        For Each t As ISqlTempTableExpression In _tempTblList
            If t.RenderType.HasFlag(TempTableRenderTypes.Drop) Then
                AppendLineOrSpace(sb)
                sb.Append(t.RenderCleanup())
            End If
        Next

        Return sb.ToString()
    End Function

#End Region

#Region "CTE"

    Private ReadOnly _cteList As New List(Of ISqlCteExpression)

    Public Function AddCte(ByVal cteExpression As ISqlCteExpression) As ISqlCteExpression Implements ISqlBuilder.AddCte
        If IsCteWithSameNameAlreadyAdded(cteExpression) Then 'to make sure we don't get multiple cte's with the same name
            Throw New Exception(String.Concat("Attempting to add a CTE with a duplicate name. Name: ", cteExpression.CteName))
        Else
            _cteList.Add(cteExpression)
        End If
        Return cteExpression
    End Function

    ''' <summary>
    ''' Adds a CTE to be rendered to the out sql.
    ''' </summary>
    ''' <typeparam name="TOtherTable">The type of table the CTE SELECT expression is built from (other table because it doesn't have to be the same type as this SqlBuilder's type)</typeparam>
    ''' <param name="baseSqlBuilder">SqlSelectBuilder corresponding to the CTE's SELECT expression</param>
    ''' <param name="cteName">The name of the CTE (;WITH name AS)</param>
    ''' <param name="alias">The alias used for the CTE in the FROM clause</param>
    ''' <param name="unionQueries">Additional 'unionable' queries to be joined with the query </param>
    ''' <returns>The CTE expression added to the SqlBuilder</returns>
    Public Function AddCte(Of TOtherTable As {ISqlBaseTableExpression, New})(ByVal baseSqlBuilder As IBTSqlSelectBuilder(Of TOtherTable), ByVal cteName As String, ByVal [alias] As String, ByVal ParamArray unionQueries As ISqlUnionable()) As ISqlCteExpression Implements ISqlBuilder.AddCte
        If unionQueries.Any Then
            Return AddCte(New BTSqlUnionBuilder(Of TOtherTable)(baseSqlBuilder, unionQueries), cteName, [alias])
        Else
            'no need to create a union builder if we don't need it
            Return AddCte(New BtSqlCteExpression(Of TOtherTable)(baseSqlBuilder, cteName, [alias]))
        End If
    End Function

    ''' <summary>
    ''' Adds a CTE to be rendered to the out sql.
    ''' </summary>
    ''' <typeparam name="TOtherTable">The type of table the CTE SELECT expression is built from (other table because it doesn't have to be the same type as this SqlBuilder's type)</typeparam>
    ''' <param name="unionBuilder">SqlUnionBuilder corresponding to the CTE's SELECT expression with all UNION's</param>
    ''' <param name="cteName">The name of the CTE (;WITH name AS)</param>
    ''' <param name="alias">The alias used for the CTE in the FROM clause</param>
    ''' <returns>The CTE expression added to the SqlBuilder</returns>
    Public Function AddCte(Of TOtherTable As {ISqlBaseTableExpression, New})(ByVal unionBuilder As IBTSqlUnionBuilder(Of TOtherTable), ByVal cteName As String, ByVal [alias] As String) As ISqlCteExpression Implements ISqlBuilder.AddCte
        Return AddCte(New BtSqlCteExpression(Of TOtherTable)(unionBuilder, cteName, [alias]))
    End Function

    Public Function AddCte(ByVal cteExpression As ISqlCteExpression, ByVal position As Integer) As ISqlCteExpression Implements ISqlBuilder.AddCte
        If IsCteWithSameNameAlreadyAdded(cteExpression) Then 'to make sure we don't get multiple cte's with the same name
            Throw New Exception(String.Concat("Attempting to add a CTE with a duplicate name. Name: ", cteExpression.CteName))
        Else
            position = If(position < 0, 0, position)
            position = If(position >= _cteList.Count, _cteList.Count - 1, position)
            _cteList.Insert(position, cteExpression)
        End If
        Return cteExpression
    End Function

    ''' <summary>
    ''' Adds a CTE to be rendered to the out sql.
    ''' </summary>
    ''' <typeparam name="TOtherTable">The type of table the CTE SELECT expression is built from (other table because it doesn't have to be the same type as this SqlBuilder's type)</typeparam>
    ''' <param name="baseSqlBuilder">SqlSelectBuilder corresponding to the CTE's SELECT expression</param>
    ''' <param name="cteName">The name of the CTE (;WITH name AS)</param>
    ''' <param name="alias">The alias used for the CTE in the FROM clause</param>
    ''' <param name="position">The position to insert the CTE</param>
    ''' <param name="unionQueries">Additional 'unionable' queries to be joined with the query </param>
    ''' <returns>The CTE expression added to the SqlBuilder</returns>
    Public Function AddCte(Of TOtherTable As {ISqlBaseTableExpression, New})(ByVal baseSqlBuilder As IBTSqlSelectBuilder(Of TOtherTable), ByVal cteName As String, ByVal [alias] As String, ByVal position As Integer, ByVal ParamArray unionQueries As ISqlUnionable()) As ISqlCteExpression Implements ISqlBuilder.AddCte
        If unionQueries.Any Then
            Return AddCte(New BTSqlUnionBuilder(Of TOtherTable)(baseSqlBuilder, unionQueries), cteName, [alias], position)
        Else
            'no need to create a union builder if we don't need it
            Return AddCte(New BtSqlCteExpression(Of TOtherTable)(baseSqlBuilder, cteName, [alias]), position)
        End If
    End Function

    ''' <summary>
    ''' Adds a CTE to be rendered to the out sql.
    ''' </summary>
    ''' <typeparam name="TOtherTable">The type of table the CTE SELECT expression is built from (other table because it doesn't have to be the same type as this SqlBuilder's type)</typeparam>
    ''' <param name="unionBuilder">SqlUnionBuilder corresponding to the CTE's SELECT expression with all UNION's</param>
    ''' <param name="cteName">The name of the CTE (;WITH name AS)</param>
    ''' <param name="alias">The alias used for the CTE in the FROM clause</param>
    ''' <param name="position">The position to insert the CTE</param>
    ''' <returns>The CTE expression added to the SqlBuilder</returns>
    Public Function AddCte(Of TOtherTable As {ISqlBaseTableExpression, New})(ByVal unionBuilder As IBTSqlUnionBuilder(Of TOtherTable), ByVal cteName As String, ByVal [alias] As String, ByVal position As Integer) As ISqlCteExpression Implements ISqlBuilder.AddCte
        Return AddCte(New BtSqlCteExpression(Of TOtherTable)(unionBuilder, cteName, [alias]), position)
    End Function

    Public Function IsCteWithSameNameAlreadyAdded(ByVal cteExpression As ISqlCteExpression) As Boolean Implements ISqlBuilder.IsCteWithSameNameAlreadyAdded
        Return _cteList.Where(Function(c) c.CteName.ToLower = cteExpression.CteName.ToLower).Any
    End Function

    Private Function GetSqlWithClause() As String
        Dim sb As New StringBuilder()

        Dim baseCte As ISqlCteExpression = TryCast(_from, ISqlCteExpression)
        Dim hadFirst As Boolean
        If baseCte IsNot Nothing AndAlso Not IsCteWithSameNameAlreadyAdded(baseCte) Then
            sb.Append(baseCte.RenderForCte(True)) 'just in case AddCte wasn't called for the From Cte, add it here first
            AppendLineOrSpace(sb)
            hadFirst = True
        End If

        For Each c As ISqlCteExpression In _cteList
            sb.Append(c.RenderForCte(Not hadFirst))
            hadFirst = True
            AppendLineOrSpace(sb)
        Next

        Return sb.ToString()
    End Function

#End Region

#Region "Join"

    Protected _joins As New List(Of ISqlJoinExpression)

    Public Function AddJoin(ByVal join As ISqlJoinExpression) As ISqlJoinExpression Implements ISqlBuilder.AddJoin
        _joins.Add(join)
        Return join
    End Function

    Public Function AddJoinOuterApply(ByVal tbl As IBTSqlGenericFromTable) As ISqlJoinExpression Implements ISqlBuilder.AddJoinOuterApply
        Return AddJoin(New BTSqlCrossApply(tbl.SubQuery, tbl.Alias))
    End Function

    Public Function AddCrossJoin(Of TJoinTable As {IBaseTableExpression})(ByVal tbl As TJoinTable) As ISqlJoinExpression Implements ISqlBuilder.AddCrossJoin
        Return AddJoin(New BTSqlTableJoin(Of TJoinTable)(tbl))
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
    ''' used in the query (the base table of the sqlbuilder, or another previously joined table).  Failure to put the columns in this order can cause issues
    ''' with alias renaming when using paging and infinite scroll options.</remarks>
    Public Function AddJoin(Of TJoinTable As {IBaseTableExpression})(ByVal joinType As JoinTypes, ByVal table As TJoinTable, ByVal column As IBTSqlColumnBase, ByVal otherColumn As ISqlSelectExpression, ByVal ParamArray filters() As ISqlFilterExpression) As ISqlJoinExpression Implements ISqlBuilder.AddJoin
        Return AddJoin(New BTSqlTableJoin(Of TJoinTable)(joinType, table, column, otherColumn, filters))
    End Function

    Public Function IsJoinAlreadyAdded(ByVal join As ISqlJoinExpression) As Boolean Implements ISqlBuilder.IsJoinAlreadyAdded
        Dim result As Boolean = False
        For Each j As ISqlJoinExpression In _joins
            If j.Alias = join.Alias Then
                result = True
                Exit For
            End If
        Next
        Return result
    End Function

    'NOTE: this function is declared as Shadows in BTSqlSelectBuilder because it requires a parameter to be provided
    Protected Function GetSqlJoinClause() As String
        Dim sb As New StringBuilder()
        For i As Integer = 0 To _joins.Count - 1
            Dim join As ISqlJoinExpression = _joins(i)
            sb.Append(join.Render())
            AppendLineOrSpace(sb)
        Next
        Return sb.ToString()
    End Function

#End Region

#Region "Filter"

    Protected _filterList As New List(Of ISqlFilterExpression)

    Public Sub AddFilter(ByVal filter As ISqlFilterExpression) Implements ISqlBuilder.AddFilter
        _filterList.Add(filter)
    End Sub

    Public Function AddFilterGroup(ByVal booleanOperator As BooleanOperatorTypes, ParamArray filters As ISqlFilterExpression()) As IBTSqlFilterGroup Implements ISqlBuilder.AddFilterGroup
        Dim g As New BTSqlFilterGroup(booleanOperator, filters)
        _filterList.Add(g)
        Return g
    End Function

    Public Sub AddFilter(ByVal leftExpression As ISqlSelectExpression, ByVal comparisonOperator As ComparisonOperatorTypes, ByVal rightClause As String, Optional ByVal booleanOperator As BooleanOperatorTypes = BooleanOperatorTypes.AND) Implements ISqlBuilder.AddFilter
        _filterList.Add(New BTSqlComparisonFilter(leftExpression, comparisonOperator, rightClause, booleanOperator))
    End Sub

    Public Sub AddFilter(ByVal leftExpression As ISqlSelectExpression, ByVal comparisonOperator As ComparisonOperatorTypes, ByVal rightExpression As ISqlSelectExpression, Optional ByVal booleanOperator As BooleanOperatorTypes = BooleanOperatorTypes.AND) Implements ISqlBuilder.AddFilter
        _filterList.Add(New BTSqlComparisonFilter(leftExpression, comparisonOperator, rightExpression, booleanOperator))
    End Sub

    Public Sub AddFilter(ByVal leftExpression As ISqlSelectExpression, ByVal logicalOpertator As LogicalOperatorTypes, ByVal rightExpression As ISqlSelectExpression, Optional ByVal booleanOperator As BooleanOperatorTypes = BooleanOperatorTypes.AND) Implements ISqlBuilder.AddFilter
        _filterList.Add(New BTSqlLogicalFilter(leftExpression, logicalOpertator, rightExpression, booleanOperator))
    End Sub

    Public Sub AddFilter(ByVal logicalOperator As LogicalOperatorTypes, ByVal rightExpression As ISqlSelectQuery, Optional ByVal booleanOperator As BooleanOperatorTypes = BooleanOperatorTypes.AND) Implements ISqlBuilder.AddFilter
        _filterList.Add(New BTSqlLogicalFilter(logicalOperator, rightExpression, booleanOperator))
    End Sub

    Public Sub AddFilter(ByVal leftExpression As ISqlSelectExpression, ByVal logicalOperator As LogicalOperatorTypes, Optional ByVal booleanOperator As BooleanOperatorTypes = BooleanOperatorTypes.AND) Implements ISqlBuilder.AddFilter
        _filterList.Add(New BTSqlLogicalFilter(leftExpression, logicalOperator, booleanOperator))
    End Sub

    Public Sub AddFilter(ByVal leftExpression As ISqlSelectExpression, ByVal logicalOperator As LogicalOperatorTypes, ByVal logicalExpression As String, Optional ByVal booleanOperator As BooleanOperatorTypes = BooleanOperatorTypes.AND) Implements ISqlBuilder.AddFilter
        _filterList.Add(New BTSqlLogicalFilter(leftExpression, logicalOperator, logicalExpression, booleanOperator))
    End Sub

    Public Sub AddFilter(ByVal leftExpression As ISqlSelectExpression, ByVal logicalOperator As LogicalOperatorTypes, ByVal leftBetween As IBTSqlParameter, ByVal rightBetween As IBTSqlParameter, Optional ByVal booleanOperator As BooleanOperatorTypes = BooleanOperatorTypes.AND) Implements ISqlBuilder.AddFilter
        _filterList.Add(New BTSqlLogicalFilter(leftExpression, logicalOperator, leftBetween, rightBetween, booleanOperator))
    End Sub

    Public Sub AddFilter(ByVal leftExpression As ISqlSelectExpression, ByVal comparisonOperator As ComparisonOperatorTypes, ByVal rightExpression As IBTSqlParameter, Optional ByVal booleanOperator As BooleanOperatorTypes = BooleanOperatorTypes.AND) Implements ISqlBuilder.AddFilter
        _filterList.Add(New BTSqlComparisonFilter(leftExpression, comparisonOperator, rightExpression, booleanOperator))
    End Sub

    Public Sub AddFilter(ByVal leftExpression As ISqlSelectExpression, ByVal logicalOperator As LogicalOperatorTypes, ByVal rightExpression As IBTSqlParameter, Optional ByVal booleanOperator As BooleanOperatorTypes = BooleanOperatorTypes.AND) Implements ISqlBuilder.AddFilter
        _filterList.Add(New BTSqlLogicalFilter(leftExpression, logicalOperator, rightExpression, booleanOperator))
    End Sub

    Public Sub AddJobsiteFilter(leftExpression As ISqlSelectExpression, filter As IJobsiteFilter) Implements ISqlBuilder.AddJobsiteFilter
        If filter.ShouldUseSubQuery Then
            AddFilter(leftExpression, LogicalOperatorTypes.In, filter.SubQuery)
        Else
            AddFilter(leftExpression, LogicalOperatorTypes.In, AddParameter("@jobFilter", filter.JobIDs))
        End If
    End Sub

    Protected Function GetSqlWhereClause() As String
        Dim sb As New StringBuilder()
        For i As Integer = 0 To _filterList.Count - 1
            Dim filter As ISqlFilterExpression = _filterList(i)
            filter.IsFirstFilter = (i = 0)
            sb.Append(filter.Render())
            AppendLineOrSpace(sb)
            filter.IsFirstFilter = False
        Next
        Return sb.ToString()
    End Function

#End Region

End Class
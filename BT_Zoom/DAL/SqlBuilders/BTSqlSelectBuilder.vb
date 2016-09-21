Option Strict On
Option Explicit On

Imports System.Collections.Generic
Imports System.Text
Imports System.Data.SqlClient
Imports System.Diagnostics
Imports System.Linq
Imports System.Text.RegularExpressions
Imports BT_Zoom.BTSql
Imports BT_Zoom.Enums.BTSql
Imports BT_Zoom.Interfaces
Imports BT_Zoom.Constants
Imports BT_Zoom.Enums.BTDebug

<Serializable>
Public Class BTSqlSelectBuilder(Of TTable As {IBaseTableExpression, New})
    Inherits BTSqlBuilderBase(Of TTable)
    Implements ISqlSelectBuilder, IBTSqlSelectBuilder(Of TTable)

    Public Sub New(ByVal fromTable As TTable, Optional ByVal useLineBreaks As Boolean = True)
        MyBase.New(fromTable, useLineBreaks)
    End Sub

    Public Sub New(Optional ByVal useLineBreaks As Boolean = True)
        MyBase.New(Activator.CreateInstance(Of TTable), useLineBreaks)
    End Sub

#Region "Constants"

    Private Const Sql_SELECT As String = "SELECT"
    Private Const Sql_GROUPBY As String = "GROUP BY"
    Private Const Sql_HAVING As String = "HAVING"
    Private Const Sql_ORDERBY As String = "ORDER BY"
    Private Const Sql_PageSubQueryAlias As String = "btSqlPage"
    Private Const Sql_PageSizeMin As Integer = 1
    Private Const Sql_PageSizeMax As Integer = 500
    Private Const Sql_DISTINCT As String = "DISTINCT"

#End Region

#Region "Select"

    Private _selectList As New List(Of ISqlSelectExpression)
    Private _selectDependentList As New List(Of ISqlSelectExpression)
    Private _selectListAppliedToPageQuery As New List(Of ISqlSelectExpression)
    Private _selectDependentListAppliedToPageQuery As New List(Of ISqlSelectExpression)

    Public ReadOnly Property SelectList As List(Of ISqlSelectExpression) Implements ISqlSelectBuilder.SelectList
        Get
            Return _selectList
        End Get
    End Property

    Public ReadOnly Property SelectDependentList As List(Of ISqlSelectExpression) Implements ISqlSelectBuilder.SelectDependentList
        Get
            Dim expressions As New List(Of ISqlSelectExpression)
            expressions.AddRange(_selectDependentList)
            expressions.AddRange(_selectDependentListAppliedToPageQuery)
            Return expressions
        End Get
    End Property

    Public ReadOnly Property SelectListAppliedToPageQuery As List(Of ISqlSelectExpression) Implements ISqlSelectBuilder.SelectListAppliedToPageQuery
        Get
            Return _selectListAppliedToPageQuery
        End Get
    End Property

    Public Sub AddSelect(ByVal ParamArray expressions() As ISqlSelectExpression) Implements ISqlSelectBuilder.AddSelect
        If expressions IsNot Nothing AndAlso expressions.Length > 0 Then
            _selectList.AddRange(expressions)
        End If
    End Sub
    Public Sub AddSelect(ByVal fnType As FunctionTypes, ByVal [alias] As String, ByVal ParamArray dependencies() As ISqlSelectExpression) Implements ISqlSelectBuilder.AddSelect
        AddSelect(New BTSqlFunctionExpression(fnType, [alias], dependencies))
    End Sub
    Public Sub AddSelect(ByVal fnType As FunctionTypes, ByVal ParamArray dependencies() As ISqlSelectExpression) Implements ISqlSelectBuilder.AddSelect
        AddSelect(New BTSqlFunctionExpression(fnType, dependencies))
    End Sub
    Public Sub AddSelect(ByVal fn As IBTSqlFunction, ByVal [alias] As String, ByVal ParamArray dependencies() As ISqlSelectExpression) Implements ISqlSelectBuilder.AddSelect
        AddSelect(New BTSqlFunctionExpression(fn, [alias], dependencies))
    End Sub
    Public Sub AddSelect(ByVal fn As IBTSqlFunction, ByVal ParamArray dependencies() As ISqlSelectExpression) Implements ISqlSelectBuilder.AddSelect
        AddSelect(New BTSqlFunctionExpression(fn, dependencies))
    End Sub
    Public Sub AddTransform(ByVal newColumnTransform As IBTSqlTransformBase) Implements ISqlSelectBuilder.AddTransform
        For Each expression As ISqlSelectExpression In newColumnTransform.ExpressionList
            AddSelectDependentIfNecessary(expression)
        Next
    End Sub

    Private Sub AddSelectDependentIfNecessary(ByVal expression As ISqlSelectExpression)
        If _selectList.FindIndex(Function(x As ISqlSelectExpression) x.GetDataRowColumnName() = expression.GetDataRowColumnName()) < 0 AndAlso
            _selectDependentList.FindIndex(Function(x As ISqlSelectExpression) x.GetDataRowColumnName() = expression.GetDataRowColumnName()) < 0 Then
            _selectDependentList.Add(expression)
        End If
    End Sub
    Private Function GetSqlSelectClause() As String
        Dim sb As New StringBuilder()
        Dim selectListCombined As New List(Of ISqlSelectExpression)
        selectListCombined.AddRange(_selectList)
        selectListCombined.AddRange(_selectDependentList)
        For i As Integer = 0 To selectListCombined.Count - 1
            Dim expression As ISqlSelectExpression = selectListCombined(i)
            If i > 0 Then
                sb.Append(", ")
            End If
            sb.Append(expression.Render())
        Next
        Return sb.ToString()
    End Function
    Private Function GetSqlSelectClauseAppliedToPageQuery() As String
        Dim sb As New StringBuilder()
        Dim expressions As New List(Of ISqlSelectExpression)
        expressions.AddRange(_selectListAppliedToPageQuery)
        expressions.AddRange(_selectDependentListAppliedToPageQuery)
        For i As Integer = 0 To expressions.Count - 1
            Dim expression As ISqlSelectExpression = expressions(i)
            If i > 0 Then
                sb.Append(", ")
            End If
            sb.Append(expression.Render())
        Next
        Return sb.ToString()
    End Function

#End Region

#Region "From"

    Public ReadOnly Property From As TTable Implements IBTSqlSelectBuilder(Of TTable).From
        Get
            Return _from
        End Get
    End Property

    Private Function GetSqlFromClause() As String
        Return From.Render()
    End Function

#End Region

#Region "Join"

    Private _joinsAppliedToPageQuery As New List(Of ISqlJoinExpression)

    Protected Shadows Function GetSqlJoinClause(ByVal dependencyIdentifiers As List(Of String)) As String
        Dim sb As New StringBuilder()
        For i As Integer = 0 To _joins.Count - 1
            Dim join As ISqlJoinExpression = _joins(i)
            If dependencyIdentifiers.Contains(join.Alias) OrElse join.ShouldNotRemove Then
                sb.Append(join.Render())
                AppendLineOrSpace(sb)
            End If
        Next
        Return sb.ToString()
    End Function

    Private Function GetSqlJoinClauseAppliedToPageQuery(ByVal dependencyIdentifiers As List(Of String)) As String
        Dim sb As New StringBuilder()
        For i As Integer = 0 To _joinsAppliedToPageQuery.Count - 1
            Dim join As ISqlJoinExpression = _joinsAppliedToPageQuery(i)
            sb.Append(join.Render())
            AppendLineOrSpace(sb)
        Next
        Return sb.ToString()
    End Function

#End Region

#Region "Filter"

    Private _filterListAppliedToPageQuery As New List(Of ISqlFilterExpression)

    Private Function GetSqlWhereClauseAppliedToPageQuery() As String
        Dim sb As New StringBuilder()
        For i As Integer = 0 To _filterListAppliedToPageQuery.Count - 1
            Dim filter As ISqlFilterExpression = _filterListAppliedToPageQuery(i)
            sb.Append(filter.Render())
            AppendLineOrSpace(sb)
        Next
        Return sb.ToString()
    End Function

    ''' <summary>
    ''' Adds LIKE filters for the provided searchFields
    ''' </summary>
    ''' <remarks>Moved from BTBaseList.AddKeywordSearchFilter</remarks>
    Public Function AddKeywordSearchFilter(parameterName As String,
                                           leftExpressions As List(Of ISqlSelectExpression),
                                           rightExpressions As List(Of String),
                                           Optional booleanOperator As BooleanOperatorTypes = BooleanOperatorTypes.AND) As IBTSqlFilterGroup Implements ISqlSelectBuilder.AddKeywordSearchFilter
        Return AddFilterGroup(booleanOperator, GenerateKeywordSearchFilterGroup(parameterName, leftExpressions, rightExpressions, booleanOperator))
    End Function

    ''' <summary>
    ''' Creates a Single Keyword Filter
    ''' </summary>
    ''' <param name="leftExpression"></param>
    ''' <param name="parameter"></param>
    ''' <remarks></remarks>
    Public Sub AddKeywordSearchFilter(leftExpression As ISqlSelectExpression,
                                            parameter As IBTSqlParameter) Implements ISqlSelectBuilder.AddKeywordSearchFilter
        AddFilter(New BTSqlLogicalFilter(leftExpression, LogicalOperatorTypes.Like, String.Format("%'+{0}+'%", parameter.Parameter.ParameterName)))
    End Sub

    ''' <summary>
    ''' Creates filters for the provided searchFields. Does not add filter group to base list query.
    ''' </summary>
    ''' <remarks></remarks>
    Public Function GenerateKeywordSearchFilterGroup(ByVal parameterName As String, ByVal leftExpressions As List(Of ISqlSelectExpression), ByVal rightExpressions As List(Of String), Optional ByVal booleanOperator As BooleanOperatorTypes = BooleanOperatorTypes.AND, Optional ByVal partialWordSearch As Boolean = False, Optional fullTextSearch As Boolean = False) As IBTSqlFilterGroup Implements ISqlSelectBuilder.GenerateKeywordSearchFilterGroup
        Dim fg As IBTSqlFilterGroup = Nothing
        Dim nonBlankExprs As IEnumerable(Of String) = rightExpressions.Where(Function(s) Not String.IsNullOrEmpty(s))
        If nonBlankExprs.Any Then
            fg = New BTSqlFilterGroup(booleanOperator)
            For i As Integer = 0 To nonBlankExprs.Count - 1
                Dim partialWordWildcard As String = If(partialWordSearch, "*", String.Empty)
                Dim pName As String = String.Format("@{0}{1}", parameterName, i)

                AddParameter(pName, If(fullTextSearch,
                                       String.Format("""{0}{1}""", nonBlankExprs(i), partialWordWildcard),
                                       nonBlankExprs(i)),
                                   If(fullTextSearch, BTConstants.KeywordFullTextSearchParameterLength, BTConstants.KeywordSearchParameterLength))

                For j As Integer = 0 To leftExpressions.Count - 1
                    If fullTextSearch Then
                        fg.AddFilter(leftExpressions(j), LogicalOperatorTypes.Contains, pName, If(j <> 0, BooleanOperatorTypes.OR, booleanOperator))
                    Else
                        fg.AddFilter(leftExpressions(j), LogicalOperatorTypes.Like, String.Format("%'+{0}+'%", pName), If(j <> 0, BooleanOperatorTypes.OR, booleanOperator))
                    End If
                Next
            Next
        End If
        Return fg
    End Function
#End Region

#Region "Group By"

    Private _groupByList As New List(Of ISqlGroupByExpression)
    Public Sub AddGroupBy(ByVal ParamArray expressions() As ISqlGroupByExpression) Implements ISqlSelectBuilder.AddGroupBy
        If expressions IsNot Nothing AndAlso expressions.Length > 0 Then
            _groupByList.AddRange(expressions)
        End If
    End Sub
    Public Sub AddGroupBy(ByVal ParamArray expressions() As ISqlSelectExpression) Implements ISqlSelectBuilder.AddGroupBy
        If expressions IsNot Nothing AndAlso expressions.Length > 0 Then
            For Each expr As ISqlSelectExpression In expressions
                _groupByList.Add(New BTSqlGroupBy(expr))
            Next
        End If
    End Sub
    Public Sub ClearSqlGroupBy() Implements ISqlSelectBuilder.ClearSqlGroupBy
        _groupByList.Clear()
    End Sub
    Private Function GetSqlGroupByClause() As String
        Dim sb As New StringBuilder()
        For i As Integer = 0 To _groupByList.Count - 1
            Dim groupBy As ISqlGroupByExpression = _groupByList(i)
            If i > 0 Then
                sb.Append(", ")
            End If
            sb.Append(groupBy.Render())
        Next
        Return sb.ToString()
    End Function

#End Region

#Region "Having"

    'NOTE: technically only aggregate functions should be used in having clauses.  For now I'm just allowing any kind of filter (function or not), and will return to this at a later date to clean it up.

    Private _havingList As New List(Of ISqlFilterExpression)
    Public Sub AddHaving(ByVal filter As ISqlFilterExpression) Implements ISqlSelectBuilder.AddHaving
        _havingList.Add(filter)
    End Sub
    Public Sub ClearSqlHaving() Implements ISqlSelectBuilder.ClearSqlHaving
        _havingList.Clear()
    End Sub
    Private Function GetSqlHavingClause() As String
        Dim sb As New StringBuilder()
        For i As Integer = 0 To _havingList.Count - 1
            Dim having As ISqlFilterExpression = _havingList(i)
            having.IsFirstFilter = (i = 0)
            sb.Append(having.Render())
            If i < _havingList.Count - 1 Then
                AppendLineOrSpace(sb)
            End If
            having.IsFirstFilter = False
        Next
        Return sb.ToString()
    End Function

#End Region

#Region "Order By"

    Private _orderByList As New List(Of ISqlOrderByExpression)
    Public Sub AddOrderBy(ByVal orderBy As ISqlOrderByExpression) Implements ISqlSelectBuilder.AddOrderBy
        _orderByList.Add(orderBy)
    End Sub
    Public Sub AddOrderBy(ByVal text As String, ByVal direction As DirectionTypes, ByVal ParamArray dependencies() As ISqlSelectExpression) Implements ISqlSelectBuilder.AddOrderBy
        AddOrderBy(New BTSqlTextOrderByExpression(text, direction, dependencies))
    End Sub
    Public Sub AddOrderBy(ByVal column As ISqlSelectExpression, ByVal direction As DirectionTypes) Implements ISqlSelectBuilder.AddOrderBy
        _orderByList.Add(New BTSqlOrderBy(column, direction))
    End Sub
    Public Sub AddOrderBy(ByVal fnType As FunctionTypes, ByVal direction As DirectionTypes, ByVal ParamArray dependencies() As ISqlSelectExpression) Implements ISqlSelectBuilder.AddOrderBy
        AddOrderBy(New BTSqlFunctionExpression(fnType, dependencies), direction)
    End Sub
    Public Sub AddOrderBy(ByVal fn As IBTSqlFunction, ByVal direction As DirectionTypes, ByVal ParamArray dependencies() As ISqlSelectExpression) Implements ISqlSelectBuilder.AddOrderBy
        AddOrderBy(New BTSqlFunctionExpression(fn, dependencies), direction)
    End Sub
    Public Sub AddOrderByAscending(ByVal column As ISqlSelectExpression) Implements ISqlSelectBuilder.AddOrderByAscending
        _orderByList.Add(New BTSqlOrderBy(column, DirectionTypes.ASC))
    End Sub
    Public Sub AddOrderByAscending(ByVal text As String, ByVal ParamArray dependencies() As ISqlSelectExpression) Implements ISqlSelectBuilder.AddOrderByAscending
        AddOrderBy(New BTSqlTextOrderByExpression(text, DirectionTypes.ASC, dependencies))
    End Sub
    Public Sub AddOrderByAscending(ByVal fnType As FunctionTypes, ByVal ParamArray dependencies() As ISqlSelectExpression) Implements ISqlSelectBuilder.AddOrderByAscending
        AddOrderByAscending(New BTSqlFunctionExpression(fnType, dependencies))
    End Sub
    Public Sub AddOrderByAscending(ByVal fn As IBTSqlFunction, ByVal ParamArray dependencies() As ISqlSelectExpression) Implements ISqlSelectBuilder.AddOrderByAscending
        AddOrderByAscending(New BTSqlFunctionExpression(fn, dependencies))
    End Sub
    Public Sub AddOrderByDescending(ByVal column As ISqlSelectExpression) Implements ISqlSelectBuilder.AddOrderByDescending
        _orderByList.Add(New BTSqlOrderBy(column, DirectionTypes.DESC))
    End Sub
    Public Sub AddOrderByDescending(ByVal text As String, ByVal ParamArray dependencies() As ISqlSelectExpression) Implements ISqlSelectBuilder.AddOrderByDescending
        AddOrderBy(New BTSqlTextOrderByExpression(text, DirectionTypes.DESC, dependencies))
    End Sub
    Public Sub AddOrderByDescending(ByVal fnType As FunctionTypes, ByVal ParamArray dependencies() As ISqlSelectExpression) Implements ISqlSelectBuilder.AddOrderByDescending
        AddOrderByDescending(New BTSqlFunctionExpression(fnType, dependencies))
    End Sub
    Public Sub AddOrderByDescending(ByVal fn As IBTSqlFunction, ByVal ParamArray dependencies() As ISqlSelectExpression) Implements ISqlSelectBuilder.AddOrderByDescending
        AddOrderByDescending(New BTSqlFunctionExpression(fn, dependencies))
    End Sub
    Public Sub ClearSqlOrderBy() Implements ISqlSelectBuilder.ClearSqlOrderBy
        _orderByList.Clear()
    End Sub
    Private Function GetSqlOrderByClause() As String
        Dim sb As New StringBuilder()
        For i As Integer = 0 To _orderByList.Count - 1
            Dim orderBy As ISqlOrderByExpression = _orderByList(i)

            'check if it's a transform
            Dim orderByColumn As ISqlOrderByColumnExpression = TryCast(orderBy, ISqlOrderByColumnExpression)
            If orderByColumn IsNot Nothing Then
                Dim transform As BTSqlTransformBase = TryCast(orderByColumn.Column, BTSqlTransformBase)
                If transform IsNot Nothing AndAlso transform.SortExpressionGroup IsNot Nothing Then
                    'overwrite the transform's primary direction of the group with the direction of this orderby clause
                    transform.SortExpressionGroup.Direction = orderByColumn.Direction
                End If
            End If

            If i > 0 Then
                sb.Append(", ")
            End If
            sb.Append(orderBy.Render())
        Next
        Return sb.ToString()
    End Function

#End Region

#Region "Dependencies"

#Region "GetDependencyIdentifiers"

    Public Function GetDependencyIdentifiers() As List(Of String) Implements ISqlSelectBuilder.GetDependencyIdentifiers
        Dim result As New List(Of String)()

        result.AddRange(GetSelectDependencyIdentifiers())

        For Each identifier As String In GetJoinDependencyIdentifiers()
            If Not result.Contains(identifier) Then
                result.Add(identifier)
            End If
        Next

        For Each identifier As String In GetFilterDependencyIdentifiers()
            If Not result.Contains(identifier) Then
                result.Add(identifier)
            End If
        Next

        For Each identifier As String In GetGroupByDependencyIdentifiers()
            If Not result.Contains(identifier) Then
                result.Add(identifier)
            End If
        Next

        For Each identifier As String In GetHavingDependencyIdentifiers()
            If Not result.Contains(identifier) Then
                result.Add(identifier)
            End If
        Next

        For Each identifier As String In GetOrderByDependencyIdentifiers()
            If Not result.Contains(identifier) Then
                result.Add(identifier)
            End If
        Next

        Return result
    End Function

    Public Function GetSelectDependencyIdentifiers() As List(Of String) Implements ISqlSelectBuilder.GetSelectDependencyIdentifiers
        Dim result As New List(Of String)()
        Dim selectListCombined As New List(Of ISqlSelectExpression)
        selectListCombined.AddRange(_selectList)
        selectListCombined.AddRange(_selectDependentList)
        selectListCombined.AddRange(_selectListAppliedToPageQuery)
        selectListCombined.AddRange(_selectDependentListAppliedToPageQuery)
        For i As Integer = 0 To selectListCombined.Count - 1
            Dim expression As ISqlSelectExpression = selectListCombined(i)
            BTSqlUtility.AddDependencyIdentifiers(result, expression)
        Next
        Return result
    End Function

    Public Function GetGroupByDependencyIdentifiers() As List(Of String) Implements ISqlSelectBuilder.GetGroupByDependencyIdentifiers
        Dim result As New List(Of String)()
        For i As Integer = 0 To _groupByList.Count - 1
            Dim expression As ISqlGroupByExpression = _groupByList(i)
            BTSqlUtility.AddDependencyIdentifiers(result, expression)
        Next
        Return result
    End Function

    Public Function GetHavingDependencyIdentifiers() As List(Of String) Implements ISqlSelectBuilder.GetHavingDependencyIdentifiers
        Dim result As New List(Of String)()
        For i As Integer = 0 To _havingList.Count - 1
            Dim expression As ISqlFilterExpression = _havingList(i)
            BTSqlUtility.AddDependencyIdentifiers(result, expression)
        Next
        Return result
    End Function

    Public Function GetOrderByDependencyIdentifiers() As List(Of String) Implements ISqlSelectBuilder.GetOrderByDependencyIdentifiers
        Dim result As New List(Of String)()
        For i As Integer = 0 To _orderByList.Count - 1
            Dim expression As ISqlOrderByExpression = _orderByList(i)
            BTSqlUtility.AddDependencyIdentifiers(result, expression)
        Next
        Return result
    End Function

    Public Function GetFilterDependencyIdentifiers() As List(Of String) Implements ISqlSelectBuilder.GetFilterDependencyIdentifiers
        Dim result As New List(Of String)()
        Dim filterListCombined As New List(Of ISqlFilterExpression)
        filterListCombined.AddRange(_filterList)
        filterListCombined.AddRange(_filterListAppliedToPageQuery)
        For i As Integer = 0 To filterListCombined.Count - 1
            Dim expression As ISqlFilterExpression = filterListCombined(i)
            BTSqlUtility.AddDependencyIdentifiers(result, expression)
        Next
        Return result
    End Function

    Public Function GetJoinDependencyIdentifiers() As List(Of String) Implements ISqlSelectBuilder.GetJoinDependencyIdentifiers
        Dim result As New List(Of String)()

        If _joins IsNot Nothing Then
            For Each j As ISqlJoinExpression In _joins
                BTSqlUtility.AddDependencyIdentifiers(result, j)
            Next
        End If
        If _joinsAppliedToPageQuery IsNot Nothing Then
            For Each j As ISqlJoinExpression In _joinsAppliedToPageQuery
                BTSqlUtility.AddDependencyIdentifiers(result, j)
            Next
        End If

        Return result
    End Function

#End Region

#Region "GetDependenciesByIdentifier"

    Public Function GetDependenciesByIdentifier(ByVal owner As String) As List(Of ISqlSelectExpression) Implements ISqlSelectBuilder.GetDependenciesByIdentifier
        Dim result As New List(Of ISqlSelectExpression)

        result.AddRange(GetSelectDependenciesByIdentifier(owner))

        For Each dependency As ISqlSelectExpression In GetJoinDependenciesByIdentifier(owner)
            result.Add(dependency)
        Next

        For Each dependency As ISqlSelectExpression In GetFilterDependenciesByIdentifier(owner)
            result.Add(dependency)
        Next

        For Each dependency As ISqlSelectExpression In GetGroupByDependenciesByIdentifier(owner)
            result.Add(dependency)
        Next

        For Each dependency As ISqlSelectExpression In GetHavingDependenciesByIdentifier(owner)
            result.Add(dependency)
        Next

        For Each dependency As ISqlSelectExpression In GetOrderByDependenciesByIdentifier(owner)
            result.Add(dependency)
        Next

        Return result
    End Function

    Public Function GetSelectDependenciesByIdentifier(ByVal owner As String) As List(Of ISqlSelectExpression) Implements ISqlSelectBuilder.GetSelectDependenciesByIdentifier
        Dim result As New List(Of ISqlSelectExpression)()
        Dim selectListCombined As New List(Of ISqlSelectExpression)
        selectListCombined.AddRange(_selectList)
        selectListCombined.AddRange(_selectDependentList)
        selectListCombined.AddRange(_selectListAppliedToPageQuery)
        selectListCombined.AddRange(_selectDependentListAppliedToPageQuery)
        For i As Integer = 0 To selectListCombined.Count - 1
            Dim expression As ISqlSelectExpression = selectListCombined(i)
            BTSqlUtility.AddDependenciesByIdentifier(result, expression, owner)
        Next
        Return result
    End Function

    Public Function GetGroupByDependenciesByIdentifier(ByVal owner As String) As List(Of ISqlSelectExpression) Implements ISqlSelectBuilder.GetGroupByDependenciesByIdentifier
        Dim result As New List(Of ISqlSelectExpression)()
        For i As Integer = 0 To _groupByList.Count - 1
            Dim expression As ISqlGroupByExpression = _groupByList(i)
            BTSqlUtility.AddDependenciesByIdentifier(result, expression, owner)
        Next
        Return result
    End Function

    Public Function GetHavingDependenciesByIdentifier(ByVal owner As String) As List(Of ISqlSelectExpression) Implements ISqlSelectBuilder.GetHavingDependenciesByIdentifier
        Dim result As New List(Of ISqlSelectExpression)()
        For i As Integer = 0 To _havingList.Count - 1
            Dim expression As ISqlFilterExpression = _havingList(i)
            BTSqlUtility.AddDependenciesByIdentifier(result, expression, owner)
        Next
        Return result
    End Function

    Public Function GetOrderByDependenciesByIdentifier(ByVal owner As String) As List(Of ISqlSelectExpression) Implements ISqlSelectBuilder.GetOrderByDependenciesByIdentifier
        Dim result As New List(Of ISqlSelectExpression)()
        For i As Integer = 0 To _orderByList.Count - 1
            Dim expression As ISqlOrderByExpression = _orderByList(i)
            BTSqlUtility.AddDependenciesByIdentifier(result, expression, owner)
        Next
        Return result
    End Function

    Public Function GetFilterDependenciesByIdentifier(ByVal owner As String) As List(Of ISqlSelectExpression) Implements ISqlSelectBuilder.GetFilterDependenciesByIdentifier
        Dim result As New List(Of ISqlSelectExpression)()
        Dim filterListCombined As New List(Of ISqlFilterExpression)
        filterListCombined.AddRange(_filterList)
        filterListCombined.AddRange(_filterListAppliedToPageQuery)
        For i As Integer = 0 To filterListCombined.Count - 1
            Dim expression As ISqlFilterExpression = filterListCombined(i)
            BTSqlUtility.AddDependenciesByIdentifier(result, expression, owner)
        Next
        Return result
    End Function

    Public Function GetJoinDependenciesByIdentifier(ByVal owner As String) As List(Of ISqlSelectExpression) Implements ISqlSelectBuilder.GetJoinDependenciesByIdentifier
        Dim result As New List(Of ISqlSelectExpression)()
        Dim joinListCombined As New List(Of ISqlJoinExpression)
        joinListCombined.AddRange(_joins)
        joinListCombined.AddRange(_joinsAppliedToPageQuery)
        For i As Integer = 0 To joinListCombined.Count - 1
            Dim expression As IHasDependencies = joinListCombined(i)
            BTSqlUtility.AddDependenciesByIdentifier(result, expression, owner)
        Next
        Return result
    End Function

#End Region

#End Region

#Region "BuildSql"

    Private _pageStart As Integer? = Nothing
    Private _pageEnd As Integer? = Nothing
    Private _entityIdExpression As ISqlSelectExpression

    Public Property IsCount As Boolean = False Implements ISqlSelectBuilder.IsCount
    Public Property IsDistinct As Boolean = False Implements ISqlSelectBuilder.IsDistinct

    <NonSerialized>
    Private _ParametersNeededByBuilder As List(Of IBTSqlParameter) = New List(Of IBTSqlParameter)
    Public Property ParametersNeededByBuilder As List(Of IBTSqlParameter) Implements ISqlSelectBuilder.ParametersNeededByBuilder
        Get
            Return _ParametersNeededByBuilder
        End Get
        Set(value As List(Of IBTSqlParameter))
            _ParametersNeededByBuilder = value
        End Set
    End Property
    Public Property EntityId As Integer? Implements ISqlSelectBuilder.EntityId
    Public Property PagingSortDirection As DirectionTypes = DirectionTypes.ASC Implements ISqlSelectBuilder.PagingSortDirection

    Public Property EntityIdExpression As ISqlSelectExpression Implements ISqlSelectBuilder.EntityIdExpression
        Get
            Return _entityIdExpression
        End Get
        Set(value As ISqlSelectExpression)
            If value Is Nothing Then
                _entityIdExpression = Nothing
            Else
                _entityIdExpression = DirectCast(value.Clone, ISqlSelectExpression)
            End If
        End Set
    End Property

    ''' <summary>
    ''' If true, suppress no filters error (i.e. Will not throw an error if the where clause is empty), only update this if you are sure that's what you need
    ''' </summary>
    Public Property AllowNoFilters As Boolean = True Implements ISqlSelectBuilder.AllowNoFilters

    ''' <summary>
    ''' If true, adds OPTION(RECOMPILE) to the end of the entire query to give it a query plan instead of having it do parameter sniffing to pick up a plan
    ''' </summary>
    Public Property EnableOptionRecompile As Boolean = False Implements ISqlSelectBuilder.EnableOptionRecompile

    ''' <summary>
    ''' If true, adds OPTION(OPTIMIZE FOR UNKNOWN) to the end of the entire query. In this case the optimizer will look at all available statistical data to reach a determination of what the values of the local variables used to generate the queryplan should be.
    ''' </summary>
    Public Property EnableOptionOptimizeForUnknown As Boolean = False Implements ISqlSelectBuilder.EnableOptionOptimizeForUnknown

    ''' <summary>
    ''' If true, adds OPTION(force order) to the end of the entire query.
    ''' </summary>
    Public Property EnableOptionForceOrder As Boolean = False Implements ISqlSelectBuilder.EnableOptionForceOrder

    Public Sub SetPagingByNumberAndSize(ByVal pageNumber As Integer?, ByVal pageSize As Integer?) Implements ISqlSelectBuilder.SetPagingByNumberAndSize
        If Not (pageNumber.HasValue AndAlso pageSize.HasValue) Then
            _pageStart = Nothing
            _pageEnd = Nothing
            Exit Sub
        End If

        If pageNumber < 1 Then
            Throw New ArgumentOutOfRangeException("pageNumber", pageNumber, "pageNumber > 0")
        End If
        If pageSize < Sql_PageSizeMin OrElse pageSize > Sql_PageSizeMax Then
            Throw New ArgumentOutOfRangeException("pageSize", pageSize, String.Format("{0} <= pageSize <= {1}", Sql_PageSizeMin, Sql_PageSizeMax))
        End If

        Dim rowIdStart As Integer = ((pageNumber.Value - 1) * pageSize.Value) + 1
        _pageStart = rowIdStart
        _pageEnd = rowIdStart + pageSize - 1
    End Sub

    Public Sub SetPagingByRows(ByVal rowNumStart As Integer?, ByVal rowNumEnd As Integer?) Implements ISqlSelectBuilder.SetPagingByRows
        If Not (rowNumStart.HasValue AndAlso rowNumStart.HasValue) Then
            _pageStart = Nothing
            _pageEnd = Nothing
            Exit Sub
        End If

        If rowNumStart.Value < 1 Then
            Throw New ArgumentOutOfRangeException("rowNumStart", rowNumStart, "rowNumStart > 0")
        End If
        If rowNumEnd.Value < rowNumStart Then
            Throw New ArgumentOutOfRangeException("rowNumEnd", rowNumStart, "rowNumEnd > rowNumStart")
        End If

        _pageStart = rowNumStart
        _pageEnd = rowNumEnd
    End Sub

    Private Sub BuildSql(ByRef sb As StringBuilder)
        If IsCount AndAlso (TopNum.HasValue OrElse _pageStart.HasValue OrElse _pageEnd.HasValue OrElse EntityId.HasValue OrElse EntityIdExpression IsNot Nothing) Then
            Throw New BTSqlException("BuildSql if isCount is TRUE then cannot specify topNum, paging or row number args")
        End If

        If TopNum.HasValue AndAlso (_pageStart.HasValue OrElse _pageEnd.HasValue OrElse EntityId.HasValue OrElse EntityIdExpression IsNot Nothing) Then
            Throw New BTSqlException("BuildSql either specify topNum OR paging OR row number")
        End If

        If (_pageStart.HasValue AndAlso Not _pageEnd.HasValue) OrElse (Not _pageStart.HasValue AndAlso _pageEnd.HasValue) Then
            Throw New BTSqlException("BuildSql must specify both (pageNumber AND pageSize) OR (rowNumStart AND rowNumEnd) if either one is provided")
        End If

        If (EntityId.HasValue AndAlso EntityIdExpression Is Nothing) OrElse (Not EntityId.HasValue AndAlso EntityIdExpression IsNot Nothing) Then
            Throw New BTSqlException("BuildSql must specify both EntityId and EntityIdExpression if either one is provided")
        End If

        Dim isPaging As Boolean = (_pageStart.HasValue AndAlso _pageEnd.HasValue)
        Dim isRowNumOver As Boolean = (EntityId.HasValue AndAlso EntityIdExpression IsNot Nothing)

        If isPaging AndAlso isRowNumOver Then
            Throw New BTSqlException("BuildSql should specify paging OR row number args, not both")
        End If

        If isRowNumOver Then
            AddSelectDependentIfNecessary(EntityIdExpression)
        End If

        If isPaging OrElse isRowNumOver Then
            OptimizeForPaging()
        End If

        Dim dependencyIdentifiers As List(Of String) = GetDependencyIdentifiers()
        Dim sqlSelect As String = GetSqlSelectClause()
        Dim sqlFrom As String = GetSqlFromClause()
        Dim sqlJoin As String = GetSqlJoinClause(dependencyIdentifiers)
        Dim sqlWhere As String = GetSqlWhereClause()
        Dim sqlGroupBy As String = GetSqlGroupByClause()
        Dim sqlHaving As String = GetSqlHavingClause()
        Dim sqlOrderBy As String = GetSqlOrderByClause()

        If isPaging OrElse isRowNumOver Then
            sb.AppendFormat("{0} {1}.*", Sql_SELECT, Sql_PageSubQueryAlias)
            Dim sqlSelectPaging As String = GetSqlSelectClauseAppliedToPageQuery()
            If Not String.IsNullOrWhitespace(sqlSelectPaging) Then
                sb.AppendFormat(", {0}", sqlSelectPaging)
            End If
            AppendLineOrSpace(sb)
            sb.AppendFormat("{0} (", Sql_FROM)
            AppendLineOrSpace(sb)
        End If

        If Not (IsCount OrElse isRowNumOver) AndAlso String.IsNullOrWhitespace(sqlSelect) Then
            Throw New BTSqlException(String.Format("{0} clause is empty", Sql_SELECT))
        End If
        sb.Append(Sql_SELECT)

        If IsDistinct Then
            sb.AppendFormat(" {0}", Sql_DISTINCT)
        End If

        If IsCount Then
            sb.AppendFormat(" {0}", New BTSqlFunctionExpression(FunctionTypes.Count))
        Else
            If TopNum.HasValue Then
                If TopNum.Value <= 0 OrElse TopNum.Value > Sql_TopMax Then
                    Throw New ArgumentOutOfRangeException("TopNum", TopNum.Value, String.Format("0 <= TopNum <= {0}", Sql_TopMax))
                End If
                sb.AppendFormat(" {0} {1}", Sql_TOP, TopNum.Value.ToString())
            ElseIf isPaging OrElse isRowNumOver Then
                If String.IsNullOrWhitespace(sqlOrderBy) Then
                    Throw New BTSqlException(String.Format("{0} clause is empty and is required for paging/rowNumber", Sql_ORDERBY))
                End If
                sb.AppendFormat(" Row_Number() OVER({0} {1}) AS {2}, COUNT(1) OVER () AS {3},", Sql_ORDERBY, sqlOrderBy, BTConstants.DAL.PagingRowIdColumnName, BTConstants.DAL.PagingTotalRowsColumnName)
            End If
            sb.AppendFormat(" {0}", sqlSelect)
        End If
        AppendLineOrSpace(sb)

        If String.IsNullOrWhitespace(sqlFrom) Then
            Throw New BTSqlException(String.Format("{0} clause is empty", Sql_FROM))
        End If
        sb.AppendFormat("{0} {1}", Sql_FROM, sqlFrom)
        AppendLineOrSpace(sb)

        If Not String.IsNullOrWhitespace(sqlJoin) Then
            sb.Append(sqlJoin)
        End If

        If String.IsNullOrWhitespace(sqlWhere) AndAlso Not AllowNoFilters Then
            Throw New BTSqlException(String.Format("{0} clause is empty", Sql_WHERE))
        End If
        If Not String.IsNullOrWhitespace(sqlWhere) Then
            sb.AppendFormat("{0} {1}", Sql_WHERE, sqlWhere)
        End If

        If Not String.IsNullOrWhitespace(sqlGroupBy) Then
            sb.AppendFormat("{0} {1}", Sql_GROUPBY, sqlGroupBy)
            AppendLineOrSpace(sb)
        End If

        If Not String.IsNullOrWhitespace(sqlHaving) Then
            sb.AppendFormat("{0} {1}", Sql_HAVING, sqlHaving)
            AppendLineOrSpace(sb)
        End If

        If isPaging OrElse isRowNumOver Then
            sb.AppendFormat(") as {0}", Sql_PageSubQueryAlias)
            Dim sqlJoinPaging As String = GetSqlJoinClauseAppliedToPageQuery(dependencyIdentifiers)
            If Not String.IsNullOrWhitespace(sqlJoinPaging) Then
                AppendLineOrSpace(sb)
                sb.Append(sqlJoinPaging)
            Else
                AppendLineOrSpace(sb)
            End If

            If isPaging Then
                'add the parameters needed for the paging query
                Dim pStart As IBTSqlParameter = DataAccessHandler.CreateSqlParameter("@pageStart", _pageStart.Value).ToBTSqlParameter
                Dim pEnd As IBTSqlParameter = DataAccessHandler.CreateSqlParameter("@pageEnd", _pageEnd.Value).ToBTSqlParameter
                ParametersNeededByBuilder.AddRange({pStart, pEnd})
                'Get the paging window by finding corresponding rowids
                sb.AppendFormat("{0} {1}.{2} {3} {4} AND {5}", Sql_WHERE, Sql_PageSubQueryAlias, BTConstants.DAL.PagingRowIdColumnName, LogicalOperatorTypes.Between.GetDescription, pStart.Parameter.ParameterName, pEnd)
            Else
                'add the parameter needed to get the correct row back
                Dim pEntityId As IBTSqlParameter = DataAccessHandler.CreateSqlParameter("@entityId", EntityId.Value).ToBTSqlParameter
                ParametersNeededByBuilder.Add(pEntityId)
                'Get the correct row by finding the corresponding entityId
                Dim f As New BTSqlComparisonFilter(EntityIdExpression, ComparisonOperatorTypes.Equals, pEntityId)
                f.UpdateOwnerForPaging(Sql_PageSubQueryAlias, EntityIdExpression.GetDependencyIdentifiers)
                f.IsFirstFilter = True
                sb.AppendFormat("{0} {1}", Sql_WHERE, f.Render())
            End If

            Dim sqlWherePaging As String = GetSqlWhereClauseAppliedToPageQuery()
            If Not String.IsNullOrWhitespace(sqlWherePaging) Then
                AppendLineOrSpace(sb)
                sb.Append(sqlWherePaging)
            End If
            AppendLineOrSpace(sb)
            sb.AppendFormat("{0} {1}.{2} {3}", Sql_ORDERBY, Sql_PageSubQueryAlias, BTConstants.DAL.PagingRowIdColumnName, PagingSortDirection.GetDescription()) 'ORDER by the row id to be sure the results are sorted correctly
            AppendLineOrSpace(sb)
        Else
            If Not String.IsNullOrWhitespace(sqlOrderBy) AndAlso Not IsCount Then
                sb.AppendFormat("{0} {1}", Sql_ORDERBY, sqlOrderBy)
                AppendLineOrSpace(sb)
            End If
        End If

        Dim isOptimizingForUnknown As Boolean = False
        If EnableOptionOptimizeForUnknown Then
            sb.Append(" OPTION(OPTIMIZE FOR UNKNOWN)")
            isOptimizingForUnknown = True
        End If

        'do not use option recompile if already optimizing for unknown!
        If Not isOptimizingForUnknown AndAlso EnableOptionRecompile Then
            sb.Append(" OPTION(RECOMPILE)")
        End If

        If EnableOptionForceOrder Then
            sb.Append(" OPTION(FORCE ORDER)")
        End If


        'for debugging
        Dim sb2 As New StringBuilder()
        sb2.AppendLine("--------------------------")
        sb2.AppendLine(sb.ToString())
        sb2.AppendLine("--------------------------")
        BTDebug.WriteLine(BTDebugOutputTypes.Sql, sb2.ToString())
    End Sub

    Private Function OptimizeForPaging() As List(Of String)
        Dim result As New List(Of String)

        Dim requiredIdentifiers As New List(Of String)()

        _joins = BTSqlUtility.ResetJoins(_joins, _joinsAppliedToPageQuery)
        _joinsAppliedToPageQuery.Clear()
        _selectDependentListAppliedToPageQuery.Clear()

        requiredIdentifiers.AddRange(GetFilterDependencyIdentifiers())
        If Not requiredIdentifiers.Contains(From.Alias) Then
            requiredIdentifiers.Add(From.Alias)
        End If

        For Each identifier As String In GetOrderByDependencyIdentifiers()
            If Not requiredIdentifiers.Contains(identifier) Then
                requiredIdentifiers.Add(identifier)
            End If
        Next
        For Each identifier As String In GetGroupByDependencyIdentifiers()
            If Not requiredIdentifiers.Contains(identifier) Then
                requiredIdentifiers.Add(identifier)
            End If
        Next
        For Each identifier As String In GetHavingDependencyIdentifiers()
            If Not requiredIdentifiers.Contains(identifier) Then
                requiredIdentifiers.Add(identifier)
            End If
        Next

        'these are all of the joins that do not need to be in the main query... we can pull them out into the paging query
        Dim pagingJoinIdentifiers As New List(Of String)
        For Each j As ISqlJoinExpression In _joins
            If CanMoveJoin(j, requiredIdentifiers, New List(Of ISqlJoinExpression) From {j}) Then
                If Not pagingJoinIdentifiers.Contains(j.Alias) Then
                    pagingJoinIdentifiers.Add(j.Alias)
                End If
            End If
        Next

        Dim pagingSelectIdentifiers As List(Of String) = GetSelectDependencyIdentifiers()

        Dim joinDependencyIdentifiers As List(Of String) = GetJoinDependencyIdentifiers()

        For i As Integer = _joins.Count - 1 To 0 Step -1
            Dim join As ISqlJoinExpression = _joins(i)
            If pagingJoinIdentifiers.Contains(join.Alias) Then
                If pagingSelectIdentifiers.Contains(join.Alias) OrElse joinDependencyIdentifiers.Contains(join.Alias) Then
                    If join.OtherColumn IsNot Nothing Then
                        Dim otherColumnDependencyIdentifiers As List(Of String) = join.OtherColumn.GetDependencyIdentifiers()
                    End If
                    join.UpdateOwnerForPaging(Sql_PageSubQueryAlias, requiredIdentifiers)
                    _joinsAppliedToPageQuery.Add(join)
                End If
                _joins.RemoveAt(i)
            End If
        Next
        _joinsAppliedToPageQuery.Reverse()

        Dim newDependencies As New List(Of ISqlSelectExpression)()

        For i As Integer = _selectDependentList.Count - 1 To 0 Step -1
            Dim selectExpression As ISqlSelectExpression = _selectDependentList(i)
            Dim canMoveIdentifiers As New List(Of String)()
            Dim cannotMoveIdentifiers As New List(Of String)()
            Dim selectDependencies As List(Of String) = selectExpression.GetDependencyIdentifiers()
            For Each identifier As String In selectDependencies
                If Not pagingJoinIdentifiers.Contains(identifier) Then
                    cannotMoveIdentifiers.Add(identifier)
                Else
                    canMoveIdentifiers.Add(identifier)
                End If
            Next
            If canMoveIdentifiers.Count > 0 Then
                If cannotMoveIdentifiers.Count > 0 Then
                    For Each identifier As String In cannotMoveIdentifiers
                        BTSqlUtility.AddDependenciesByIdentifier(newDependencies, selectExpression, identifier, True)
                    Next
                End If

                If Not _selectDependentListAppliedToPageQuery.Contains(selectExpression) Then
                    selectExpression.UpdateOwnerForPaging(Sql_PageSubQueryAlias, requiredIdentifiers)
                    _selectDependentListAppliedToPageQuery.Add(selectExpression)
                End If
                _selectDependentList.RemoveAt(i)
            End If
        Next
        _selectDependentListAppliedToPageQuery.Reverse()
        _selectDependentList.AddRange(newDependencies)
        newDependencies.Clear()

        For i As Integer = _selectList.Count - 1 To 0 Step -1
            Dim selectExpression As ISqlSelectExpression = _selectList(i)
            Dim canMoveIdentifiers As New List(Of String)()
            Dim cannotMoveIdentifiers As New List(Of String)()
            Dim selectDependencies As List(Of String) = selectExpression.GetDependencyIdentifiers()
            For Each identifier As String In selectDependencies
                If Not pagingJoinIdentifiers.Contains(identifier) Then
                    cannotMoveIdentifiers.Add(identifier)
                Else
                    canMoveIdentifiers.Add(identifier)
                End If
            Next
            If canMoveIdentifiers.Count > 0 Then
                If cannotMoveIdentifiers.Count > 0 Then
                    For Each identifier As String In cannotMoveIdentifiers
                        BTSqlUtility.AddDependenciesByIdentifier(newDependencies, selectExpression, identifier, True)
                    Next
                End If
                selectExpression.UpdateOwnerForPaging(Sql_PageSubQueryAlias, requiredIdentifiers)
                _selectListAppliedToPageQuery.Add(selectExpression)
                _selectList.RemoveAt(i)
            End If
        Next
        _selectListAppliedToPageQuery.Reverse()
        _selectList.AddRange(newDependencies)
        newDependencies.Clear()

        For i As Integer = _filterList.Count - 1 To 0 Step -1
            Dim filterExpression As ISqlFilterExpression = _filterList(i)
            Dim canMove As Boolean = True
            Dim filterDependencies As List(Of String) = filterExpression.GetDependencyIdentifiers()
            For Each identifier As String In filterDependencies
                If Not pagingJoinIdentifiers.Contains(identifier) Then
                    canMove = False
                    Exit For
                End If
            Next
            If canMove Then
                filterExpression.UpdateOwnerForPaging(Sql_PageSubQueryAlias, requiredIdentifiers)
                _filterListAppliedToPageQuery.Add(filterExpression)
                _filterList.RemoveAt(i)
            End If
        Next
        _filterListAppliedToPageQuery.Reverse()

        Return result
    End Function

    Private Function CanMoveJoin(ByVal join As ISqlJoinExpression, ByVal requiredIdentifiers As List(Of String), ByVal doNotProcessJoins As List(Of ISqlJoinExpression)) As Boolean
        If join.ShouldNotRemove Then 'if it's marked to not remove, it's probably being used to filter the results and therefore has an affect on the paging window
            Return False
        End If

        If requiredIdentifiers.Contains(join.Alias) Then
            Return False
        End If

        Dim joinsThatIDependOn As New List(Of ISqlJoinExpression)
        GetAllOfTheJoinsThatIDependOn(join, joinsThatIDependOn)

        Dim doNotProcessForIDependOn As Dictionary(Of String, ISqlJoinExpression) = doNotProcessJoins.ToDictionary(Function(sqlj) sqlj.Alias)
        For Each j As ISqlJoinExpression In joinsThatIDependOn
            If Not doNotProcessForIDependOn.ContainsKey(j.Alias) Then
                Dim dnp As New List(Of ISqlJoinExpression)(doNotProcessForIDependOn.Values)
                dnp.Add(j)
                If Not CanMoveJoin(j, requiredIdentifiers, dnp) Then
                    Return False
                End If
            End If
        Next

        Dim joinsThatDependOnMe As New List(Of ISqlJoinExpression)
        GetAllOfTheJoinsThatDependOnMe(join, joinsThatDependOnMe)

        Dim doNotProcessForDependOnMe As Dictionary(Of String, ISqlJoinExpression) = doNotProcessJoins.ToDictionary(Function(sqlj) sqlj.Alias)
        For Each j As ISqlJoinExpression In joinsThatDependOnMe
            If Not doNotProcessForDependOnMe.ContainsKey(j.Alias) Then
                Dim dnp As New List(Of ISqlJoinExpression)(doNotProcessForDependOnMe.Values)
                dnp.Add(j)
                If Not CanMoveJoin(j, requiredIdentifiers, dnp) Then
                    Return False
                End If
            End If
        Next

        Return True

    End Function

    Private Sub GetAllOfTheJoinsThatIDependOn(ByVal join As ISqlJoinExpression, ByVal dependencies As List(Of ISqlJoinExpression))
        If join.OtherColumn IsNot Nothing Then
            Dim joinDependencyIdentifiers As List(Of String) = join.OtherColumn.GetDependencyIdentifiers()
            For Each id As String In joinDependencyIdentifiers
                Dim otherJoin As ISqlJoinExpression = _joins.FirstOrDefault(Function(x) x.Alias = id)
                If otherJoin IsNot Nothing Then
                    If dependencies.FindIndex(Function(x) x.Alias = id) < 0 Then
                        dependencies.Add(otherJoin)
                        GetAllOfTheJoinsThatIDependOn(otherJoin, dependencies)
                    End If
                End If
            Next

        End If
    End Sub

    Private Sub GetAllOfTheJoinsThatDependOnMe(ByVal join As ISqlJoinExpression, ByVal dependencies As List(Of ISqlJoinExpression))
        For Each j As ISqlJoinExpression In _joins
            If j.Alias = join.Alias Then
                Continue For
            End If
            If j.OtherColumn IsNot Nothing Then
                Dim joinDependencyIdentifiers As List(Of String) = j.OtherColumn.GetDependencyIdentifiers()
                If joinDependencyIdentifiers.Contains(join.Alias) Then
                    If dependencies.FindIndex(Function(x) x.Alias = j.Alias) < 0 Then
                        dependencies.Add(j)
                        GetAllOfTheJoinsThatDependOnMe(j, dependencies)
                    End If
                End If
            End If
        Next
    End Sub

    ''' <summary>
    ''' Do everything that the base builder didn't already take care of...
    ''' </summary>
    Public Overrides Sub Render(ByRef sb As StringBuilder)
        BuildSql(sb)
    End Sub

    ''' <summary>
    ''' Don't render any CTE's, they've already been rendered by the SqlSelectBuilder I'm a filter in
    ''' </summary>
    ''' <remarks>As (if) we move more things out of this builder and into the base builder, we may need to change this so it renders everything but the CTE's again.</remarks>
    Public Function RenderForFilter() As String Implements ISqlFilterable.RenderForFilter
        Dim sb As New StringBuilder()
        BuildSql(sb)
        Return sb.ToString
    End Function

#End Region

#Region "Miscellaneous"

    Private Function SqlReplaceTableAliases(ByVal str As String, ByVal newTableAlias As String) As String
        Return Regex.Replace(str, "[a-zA-Z0-9]*\.", String.Format("{0}.", newTableAlias))
    End Function

    Public Overridable Sub UpdateOwnerForPaging(pagingAlias As String, ownersToChange As List(Of String)) Implements ICanUpdateOwnerForPaging.UpdateOwnerForPaging
        Dim genericTable As ICanUpdateOwnerForPaging = TryCast(From, ICanUpdateOwnerForPaging)
        If genericTable IsNot Nothing Then
            genericTable.UpdateOwnerForPaging(pagingAlias, ownersToChange)
        End If

        For Each expr As ISqlSelectExpression In _selectList
            expr.UpdateOwnerForPaging(pagingAlias, ownersToChange)
        Next

        For Each expr As ISqlSelectExpression In _selectDependentList
            expr.UpdateOwnerForPaging(pagingAlias, ownersToChange)
        Next

        For Each expr As ISqlSelectExpression In _selectListAppliedToPageQuery
            expr.UpdateOwnerForPaging(pagingAlias, ownersToChange)
        Next

        For Each expr As ISqlSelectExpression In _selectDependentListAppliedToPageQuery
            expr.UpdateOwnerForPaging(pagingAlias, ownersToChange)
        Next

        For Each expr As ISqlJoinExpression In _joins
            expr.UpdateOwnerForPaging(pagingAlias, ownersToChange)
        Next

        For Each expr As ISqlJoinExpression In _joinsAppliedToPageQuery
            expr.UpdateOwnerForPaging(pagingAlias, ownersToChange)
        Next

        For Each expr As ISqlFilterExpression In _filterList
            expr.UpdateOwnerForPaging(pagingAlias, ownersToChange)
        Next

        For Each expr As ISqlFilterExpression In _filterListAppliedToPageQuery
            expr.UpdateOwnerForPaging(pagingAlias, ownersToChange)
        Next

    End Sub

    Public Function AddAccountingIDs(fkColumn As IBTSqlColumn, tableAlias As String, columnPrefix As String) As List(Of ISqlSelectExpression) Implements ISqlSelectBuilder.AddAccountingIDs

        Dim accountingIDsTable As New BTAccountingIDsTable(tableAlias, columnPrefix)
        AddJoin(JoinTypes.LeftOuterJoin, accountingIDsTable, accountingIDsTable.AccountingID, fkColumn)
        Dim result As New List(Of ISqlSelectExpression)()
        result.AddRange(accountingIDsTable.Star())
        AddSelect(result.ToArray())
        Return result

    End Function
#End Region

End Class

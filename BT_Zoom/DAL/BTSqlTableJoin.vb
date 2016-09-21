Option Strict On
Option Explicit On

Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports BT_Zoom.Enums.BTSql
Imports BT_Zoom.Interfaces

<Diagnostics.DebuggerDisplay("Alias={Alias}")> _
<Serializable>
Public Class BTSqlTableJoin(Of TTable As {IBaseTableExpression})
    Implements ISqlJoinExpression

    ''' <summary>
    ''' Creates a new join expression which will render the join/on clause when added to a sqlbuilder
    ''' </summary>
    ''' <param name="joinType">Type of the join (e.g. Inner, Right, Left)</param>
    ''' <param name="table">Table to join with</param>
    ''' <param name="column">Column from <paramref name="table"></paramref> to be used in the ON clause</param>
    ''' <param name="otherColumn">Column from the other table we are joining with</param>
    ''' <param name="filters">Additional filters used in the ON clause</param>
    ''' <remarks>The column parameter should be a column from the TTable used as the table parameter.  The other column can be a column from any other table
    ''' used in the query (the base table of the list/sqlbuilder, or another previously joined table).  Failure to put the columns in this order can cause issues
    ''' with alias renaming when using paging and infinite scroll options.</remarks>
    Public Sub New(ByVal joinType As JoinTypes, ByVal table As TTable, ByVal column As IBTSqlColumnBase, ByVal otherColumn As ISqlSelectExpression, ByVal ParamArray filters() As ISqlFilterExpression)
        Me.JoinType = joinType
        _table = table
        _column = CType(column.WithAlias(Nothing), BTSqlColumnBase)
        _otherColumn = CType(otherColumn.WithAlias(Nothing).Clone(), ISqlSelectExpression)
        _filters = New List(Of ISqlFilterExpression)
        _filters.AddRange(filters)
    End Sub

    Public Sub New(ByVal table As TTable)
        JoinType = JoinTypes.CrossJoin
        _table = table
    End Sub

    Private _table As TTable
    Private _column As IBTSqlColumnBase
    Private _otherColumn As ISqlSelectExpression
    Private _filters As List(Of ISqlFilterExpression)

    Public ReadOnly Property [Alias] As String Implements ISqlJoinExpression.Alias
        Get
            Return _table.Alias
        End Get
    End Property

    Public Sub ChangeAlias(ByVal [alias] As String) Implements ISqlJoinable.ChangeAlias
        _table.ChangeAlias([alias])
    End Sub

    Public Property JoinType As JoinTypes Implements ISqlJoinExpression.JoinType

    ''' <summary>
    ''' When false, if the join is not needed by any dependencies (e.g. select statements, filters, etc), it will be removed from the query
    ''' </summary>
    ''' <value>Default is false (i.e. the join can be removed if the SqlBuilder wants to), consumers have to explicitly set to true in order
    ''' to not allow this to happen.  Inner joins and right joins may often be set to not allow removal as the results can differ greatly for those types of joins.</value>
    Public Property ShouldNotRemove As Boolean Implements ISqlJoinExpression.ShouldNotRemove

    Public Overrides Function ToString() As String
        If JoinType = JoinTypes.CrossJoin Then
            Return String.Format("{0} {1}", JoinType.GetDescription, _table.Render)
        End If
        'else
        Dim sb As New StringBuilder()
        sb.AppendFormat("{0} {1} ON {2} = {3}", JoinType.GetDescription(), _table.Render, _column.RenderForJoin, _otherColumn.RenderForJoin)
        For Each filter As ISqlFilterExpression In _filters
            sb.AppendFormat(" {0}", filter.Render())
        Next
        Return sb.ToString()
    End Function

    Public Function Render() As String Implements ISqlJoinExpression.Render
        Return ToString()
    End Function

    Public Sub UpdateOwnerForPaging(ByVal pagingAlias As String, ByVal ownersToChange As List(Of String)) Implements ISqlJoinExpression.UpdateOwnerForPaging
        If _otherColumn IsNot Nothing Then
            _otherColumn.UpdateOwnerForPaging(pagingAlias, ownersToChange)
        End If
        If _filters IsNot Nothing AndAlso _filters.Count > 0 Then
            For Each f As ISqlFilterExpression In _filters
                f.UpdateOwnerForPaging(pagingAlias, ownersToChange)
            Next
        End If
    End Sub

    Public ReadOnly Property OtherColumn As ISqlSelectExpression Implements ISqlJoinExpression.OtherColumn
        Get
            Return _otherColumn
        End Get
    End Property

    Public Function GetDependencyIdentifiers() As List(Of String) Implements ISqlJoinExpression.GetDependencyIdentifiers
        Dim result As New List(Of String)

        If _otherColumn IsNot Nothing Then
            BTSqlUtility.AddDependencyIdentifiers(result, _otherColumn)
        End If

        If _filters IsNot Nothing Then
            For Each f As ISqlFilterExpression In _filters
                BTSqlUtility.AddDependencyIdentifiers(result, f)
            Next
        End If

        Return result
    End Function

    Public Function GetDependenciesByIdentifier(ByVal owner As String) As List(Of ISqlSelectExpression) Implements ISqlJoinExpression.GetDependenciesByIdentifier
        Dim result As New List(Of ISqlSelectExpression)

        If _otherColumn IsNot Nothing Then
            BTSqlUtility.AddDependenciesByIdentifier(result, _otherColumn, owner)
        End If

        If _filters IsNot Nothing Then
            For Each f As ISqlFilterExpression In _filters
                BTSqlUtility.AddDependenciesByIdentifier(result, f, owner)
            Next
        End If

        Return result
    End Function

End Class

Public Class BtSqlCteJoin(Of TTable As {ISqlBaseTableExpression, New})
    Inherits BTSqlTableJoin(Of IBTSqlCteExpression(Of TTable))

    ''' <summary>
    ''' Creates a new join expression and adds it to the sql builder to be rendered
    ''' </summary>
    ''' <param name="joinType">Type of the join (e.g. Inner, Right, Left)</param>
    ''' <param name="table">CTE Table to join with</param>
    ''' <param name="column">Column from <paramref name="table"></paramref> to be used in the ON clause</param>
    ''' <param name="otherColumn">Column from the other table or CTE we are joining with</param>
    ''' <param name="filters">Additional filters used in the ON clause</param>
    ''' <remarks>The column parameter should be a column from the TJoinTable used as the table parameter.  The other column can be a column from any other table
    ''' used in the query (the base table of the sqlbuilder, or another previously joined table).  Failure to put the columns in this order can cause issues
    ''' with alias renaming when using paging and infinite scroll options.</remarks>
    Public Sub New(ByVal joinType As JoinTypes, ByVal table As IBTSqlCteExpression(Of TTable), ByVal column As IBTSqlColumnBase, ByVal otherColumn As ISqlSelectExpression, ByVal ParamArray filters() As ISqlFilterExpression)
        MyBase.New(joinType, table, column, otherColumn, filters)
        ShouldNotRemove = True 'Default is true, since most of the time, if we are going to the trouble of creating the CTE, we probably want to keep the join.
    End Sub

End Class

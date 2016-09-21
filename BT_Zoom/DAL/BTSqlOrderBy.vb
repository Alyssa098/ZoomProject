Option Strict On
Option Explicit On

Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports BT_Zoom.Enums.BTSql
Imports BT_Zoom.Interfaces


Public Class BTSqlOrderBy
    Implements ISqlOrderByColumnExpression

    Public Property Column As ISqlSelectExpression Implements ISqlOrderByColumnExpression.Column
    Public Property Direction As DirectionTypes Implements ISqlOrderByExpression.Direction

    Public Sub New(ByVal column As ISqlSelectExpression, ByVal direction As DirectionTypes)
        Me.Column = column
        Me.Direction = direction
    End Sub

    Public Shared Function CreateASC(ByVal column As ISqlSelectExpression) As BTSqlOrderBy
        Return New BTSqlOrderBy(column, DirectionTypes.ASC)
    End Function

    Public Shared Function CreateDESC(ByVal column As ISqlSelectExpression) As BTSqlOrderBy
        Return New BTSqlOrderBy(column, DirectionTypes.DESC)
    End Function

    Public Overrides Function ToString() As String
        If TypeOf Column Is BTSqlTransformBase Then
            Return Column.RenderForOrderBy() 'this already takes care of Direction
        End If

        Return String.Format("{0} {1}", Column.RenderForOrderBy(), Direction.GetDescription())
    End Function

    Public Function Render() As String Implements ISqlOrderByExpression.Render
        Return ToString()
    End Function

    Public Function GetDependencyIdentifiers() As List(Of String) Implements ISqlOrderByExpression.GetDependencyIdentifiers
        Return Column.GetDependencyIdentifiers()
    End Function

    Public Function GetDependenciesByIdentifier(ByVal owner As String) As List(Of ISqlSelectExpression) Implements ISqlOrderByExpression.GetDependenciesByIdentifier
        Return Column.GetDependenciesByIdentifier(owner)
    End Function

End Class

Public Class BTSqlOrderByGroup
    Implements ISqlOrderByColumnExpression, IBTSqlOrderByGroup

    Public Property PrimaryOrderByColumn As ISqlSelectExpression Implements ISqlOrderByColumnExpression.Column
    Public Property AdditionalOrderByList As New List(Of ISqlOrderByExpression) Implements IBTSqlOrderByGroup.AdditionalOrderByList
    Public Property Direction As DirectionTypes Implements ISqlOrderByExpression.Direction

    Public ReadOnly Property PrimaryOrderByExpression As ISqlOrderByExpression Implements IBTSqlOrderByGroup.PrimaryOrderByExpression
        Get
            Return New BTSqlOrderBy(PrimaryOrderByColumn, Direction)
        End Get
    End Property

    Public Sub New(ByVal primaryColumn As ISqlSelectExpression, ByVal primaryDirection As DirectionTypes, ParamArray addtionalOrderByExpressions As ISqlOrderByExpression())
        PrimaryOrderByColumn = primaryColumn
        Direction = primaryDirection
        AdditionalOrderByList.AddRange(addtionalOrderByExpressions)
    End Sub

    Public Sub New(ByVal primaryColumn As ISqlSelectExpression, ByVal addtionalOrderByExpressions As ISqlOrderByExpression())
        Me.New(primaryColumn, DirectionTypes.DefaultOrMatchPrimary, addtionalOrderByExpressions)
    End Sub

    ''' <summary>
    ''' Creates a new orderby group with all additional expressions following the primary sort direction
    ''' </summary>
    ''' <param name="primaryColumn">primary sort column</param>
    ''' <param name="addtionalDefaultSortOrderByExpressions">All additional parameters will for the primary sort</param>
    Public Sub New(ByVal primaryColumn As ISqlSelectExpression, ByVal ParamArray addtionalDefaultSortOrderByExpressions As ISqlSelectExpression())
        Me.New(primaryColumn, Array.ConvertAll(addtionalDefaultSortOrderByExpressions, Function(o) New BTSqlOrderBy(o, DirectionTypes.DefaultOrMatchPrimary)))
    End Sub

    ''' <summary>
    ''' Creates a new orderby group with the first item in the list as the primary item and all additional expressions following the primary sort direction
    ''' </summary>
    Public Sub New(ByVal orderByExpressions As IEnumerable(Of ISqlSelectExpression))
        'I think skip is ok to use here since according to http://msdn.microsoft.com/en-us/library/bb358985.aspx,
        'if only 1 item is in the array, it will return an empty IEnumerable(Of ISqlSelectExpression) object
        'and if there's no items or it is nothing, we have a bigger problem and it'll throw an error when it tries to get the first item
        Me.New(orderByExpressions(0), orderByExpressions.Skip(1).ToArray)
    End Sub

    Public Sub AddAdditionalOrderBy(ByVal column As ISqlSelectExpression, Optional ByVal direction As DirectionTypes = DirectionTypes.DefaultOrMatchPrimary) Implements IBTSqlOrderByGroup.AddAdditionalOrderBy
        AdditionalOrderByList.Add(New BTSqlOrderBy(column, direction))
    End Sub

    Public Sub AddAdditionalOrderByAlwaysAscending(ByVal column As ISqlSelectExpression) Implements IBTSqlOrderByGroup.AddAdditionalOrderByAlwaysAscending
        AddAdditionalOrderBy(column, DirectionTypes.ASC)
    End Sub

    Public Sub AddAdditionalOrderByAlwaysDescending(ByVal column As ISqlSelectExpression) Implements IBTSqlOrderByGroup.AddAdditionalOrderByAlwaysDescending
        AddAdditionalOrderBy(column, DirectionTypes.DESC)
    End Sub

    Public Overrides Function ToString() As String Implements IBTSqlOrderByGroup.ToString
        Dim sb As New StringBuilder(PrimaryOrderByExpression.Render)
        For Each expr As ISqlOrderByExpression In AdditionalOrderByList
            sb.Append(", ")
            If expr.Direction = DirectionTypes.DefaultOrMatchPrimary Then
                expr.Direction = Direction 'so it'll match the primary sort
            End If
            sb.Append(expr.Render)
        Next
        Return sb.ToString
    End Function

    Public Function Render() As String Implements ISqlExpression.Render
        Return ToString()
    End Function

    Public Function GetDependencyIdentifiers() As List(Of String) Implements IHasDependencies.GetDependencyIdentifiers
        Dim result As List(Of String) = PrimaryOrderByColumn.GetDependencyIdentifiers
        For Each expression As ISqlOrderByExpression In AdditionalOrderByList
            BTSqlUtility.AddDependencyIdentifiers(result, expression)
        Next
        Return result
    End Function

    Public Function GetDependenciesByIdentifier(ByVal owner As String) As List(Of ISqlSelectExpression) Implements IHasDependencies.GetDependenciesByIdentifier
        Dim result As New List(Of ISqlSelectExpression)()
        For Each expression As ISqlOrderByExpression In AdditionalOrderByList
            BTSqlUtility.AddDependenciesByIdentifier(result, expression, owner)
        Next
        Return result
    End Function

End Class
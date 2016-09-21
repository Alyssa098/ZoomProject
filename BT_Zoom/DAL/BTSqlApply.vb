Option Strict On
Option Explicit On

Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports BT_Zoom.Enums.BTSql
Imports BT_Zoom.Interfaces

<Diagnostics.DebuggerDisplay("Alias={Alias}")> _
Public MustInherit Class BTSqlApply
    Implements ISqlJoinExpression

    Public Sub New(ByVal joinType As JoinTypes, ByVal selectQuery As ISqlSelectQuery, ByVal [alias] As String)

        If joinType <> JoinTypes.OuterApply AndAlso joinType <> JoinTypes.CrossApply Then
            Throw New BTSqlException(String.Format("{0} is not allowed with BTSqlTextApplyExpression.  Please use BTSqlTextJoinExpression instead.", joinType.GetDescription()))
        End If

        If String.IsNullOrWhiteSpace([alias]) Then
            Throw New ArgumentNullException("alias", "An alias is required")
        End If

        Me.JoinType = joinType
        _selectQuery = selectQuery
        _alias = [alias]
    End Sub

    Public Property JoinType As JoinTypes Implements ISqlJoinExpression.JoinType

    ''' <summary>
    ''' When false, if the join is not needed by any dependencies (e.g. select statements, filters, etc), it will be removed from the query
    ''' </summary>
    ''' <value>Default is false, since most of the time it is safe to remove the join when doing an outer apply, if we don't need it for a dependency.</value>
    Public Property ShouldNotRemove As Boolean Implements ISqlJoinExpression.ShouldNotRemove

    Private _selectQuery As ISqlSelectQuery

    Private _alias As String
    Public ReadOnly Property [Alias] As String Implements ISqlJoinable.Alias
        Get
            Return _alias
        End Get
    End Property

    Public Sub ChangeAlias(ByVal [alias] As String) Implements ISqlJoinable.ChangeAlias
        _alias = [alias]
    End Sub

    Public Overrides Function ToString() As String
        Dim sb As New StringBuilder()
        sb.AppendFormat("{0} ", JoinType.GetDescription())
        sb.Append("(")
        sb.Append(_selectQuery.RenderForUnion())
        sb.AppendFormat(") {0}", Me.Alias)
        Return sb.ToString()
    End Function

    Public Function Render() As String Implements ISqlExpression.Render
        Return ToString()
    End Function

    Public ReadOnly Property OtherColumn As ISqlSelectExpression Implements ISqlJoinExpression.OtherColumn
        Get
            Return Nothing
        End Get
    End Property

    Public Sub UpdateOwnerForPaging(pagingAlias As String, ownersToChange As Collections.Generic.List(Of String)) Implements ISqlJoinExpression.UpdateOwnerForPaging
        _selectQuery.UpdateOwnerForPaging(pagingAlias, ownersToChange)
    End Sub

    Public Function GetDependencyIdentifiers() As List(Of String) Implements ISqlJoinExpression.GetDependencyIdentifiers
        Return _selectQuery.GetDependencyIdentifiers()
    End Function

    Public Function GetDependenciesByIdentifier(ByVal owner As String) As List(Of ISqlSelectExpression) Implements ISqlJoinExpression.GetDependenciesByIdentifier
        Return _selectQuery.GetDependenciesByIdentifier(owner)
    End Function

End Class

Public Class BTSqlOuterApply
    Inherits BTSqlApply

    Public Sub New(ByVal selectQuery As ISqlSelectQuery, ByVal [alias] As String)
        MyBase.New(JoinTypes.OuterApply, selectQuery, [alias])

    End Sub

End Class

Public Class BTSqlCrossApply
    Inherits BTSqlApply

    Public Sub New(ByVal selectQuery As ISqlSelectQuery, ByVal [alias] As String)
        MyBase.New(JoinTypes.CrossApply, selectQuery, [alias])

    End Sub

End Class

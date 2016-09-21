Option Explicit On
Option Strict On

Imports System.Text
Imports System.Collections.Generic
Imports BT_Zoom.Enums.BTSql
Imports BT_Zoom.Interfaces

Public Class BTSqlUnionSelectQuery
    Implements ISqlSelectQuery

    Private ReadOnly _query As ISqlSelectQuery

    Public Sub New(ByVal sqlBuilder As ISqlSelectBuilder, ByVal unionType As UnionTypes)
        _query = New BTSqlSelectQuery(sqlBuilder)
        Me.UnionType = unionType
    End Sub

    Public Sub New(ByVal query As ISqlSelectQuery, ByVal unionType As UnionTypes)
        _query = query
        Me.UnionType = unionType
    End Sub

    Public Property UnionType As UnionTypes

    Public Sub UpdateOwnerForPaging(pagingAlias As String, ownersToChange As Collections.Generic.List(Of String)) Implements ICanUpdateOwnerForPaging.UpdateOwnerForPaging
        _query.UpdateOwnerForPaging(pagingAlias, ownersToChange)
    End Sub

    Public Function GetDependencyIdentifiers() As Collections.Generic.List(Of String) Implements IHasDependencies.GetDependencyIdentifiers
        Return _query.GetDependencyIdentifiers()
    End Function

    Public Function GetDependenciesByIdentifier(ByVal owner As String) As List(Of ISqlSelectExpression) Implements IHasDependencies.GetDependenciesByIdentifier
        Return _query.GetDependenciesByIdentifier(owner)
    End Function

#Region "ToString / Render"

    Public Overrides Function ToString() As String
        Dim sb As New StringBuilder()
        sb.AppendLine(UnionType.GetDescription)
        sb.AppendLine()
        sb.Append(_query.RenderForUnion())
        Return sb.ToString()
    End Function

    Public Function Render() As String Implements ISqlExpression.Render
        Return ToString()
    End Function

    Public Function RenderForFilter() As String Implements ISqlSelectQuery.RenderForFilter
        Return Render()
    End Function

    Public Function RenderForUnion() As String Implements ISqlUnionable.RenderForUnion
        Return Render()
    End Function

    Public Function RenderForAssignment() As String Implements ISqlSelectQuery.RenderForAssignment
        Throw New BTSqlException("BTSqlUnionSelectQuery cannot be rendered for assignment.")
    End Function

#End Region

    Public Function Clone() As Object Implements ICloneable.Clone
        'NOTE: this performs a shallow copy
        Return New BTSqlUnionSelectQuery(_query, UnionType)
    End Function

    Public ReadOnly Property IsTableParameterTable As Boolean Implements ISqlSelectQuery.IsTableParameterTable
        Get
            Return False
        End Get
    End Property

End Class

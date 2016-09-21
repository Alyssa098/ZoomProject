Option Strict On
Option Explicit On

Imports System.Collections.Generic
Imports BT_Zoom.Interfaces

Public Class BTSqlGroupBy
    Implements ISqlGroupByExpression

    Public Property Expression As ISqlSelectExpression

    Public Sub New(ByVal expression As ISqlSelectExpression)
        Me.Expression = expression
    End Sub

    Public Overrides Function ToString() As String
        Return Expression.RenderForGroupBy()
    End Function

    Public Function Render() As String Implements ISqlGroupByExpression.Render
        Return ToString()
    End Function

    Public Function GetDependencyIdentifiers() As List(Of String) Implements ISqlGroupByExpression.GetDependencyIdentifiers
        Return Expression.GetDependencyIdentifiers()
    End Function

    Public Function GetDependenciesByIdentifier(ByVal owner As String) As List(Of ISqlSelectExpression) Implements ISqlGroupByExpression.GetDependenciesByIdentifier
        Return Expression.GetDependenciesByIdentifier(owner)
    End Function

End Class


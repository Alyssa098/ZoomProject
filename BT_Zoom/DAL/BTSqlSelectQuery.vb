Option Strict On
Option Explicit On

Imports System.Collections.Generic
Imports System.Text
Imports System.Data.SqlClient
Imports BT_Zoom.Interfaces

Public Class BTSqlSelectQuery
    Implements ISqlSelectQuery, ISqlSelectExpression, IBTSqlSelectQuery

    Public Sub New(ByVal sqlBuilder As ISqlSelectBuilder, Optional ByVal isTableParameterTable As Boolean = False, Optional ByVal [alias] As String = "")
        _sqlBuilder = sqlBuilder
        _isTableParameterTable = isTableParameterTable
        Me.Alias = [alias]
    End Sub

    Private ReadOnly _sqlBuilder As ISqlSelectBuilder

    Public Sub UpdateOwnerForPaging(ByVal pagingAlias As String, ByVal ownersToChange As List(Of String)) Implements ISqlSelectQuery.UpdateOwnerForPaging
        _sqlBuilder.UpdateOwnerForPaging(pagingAlias, ownersToChange)
    End Sub

    Public Property [Alias]() As String Implements ISqlSelectExpression.[Alias]

    Public Function WithAlias(ByVal [alias] As String) As ISqlSelectExpression Implements ISqlSelectExpression.WithAlias
        Dim result As BTSqlSelectQuery = CType(Clone(), BTSqlSelectQuery)
        result.Alias = [alias]
        Return result
    End Function

    Public Function GetDataRowColumnName() As String Implements ISqlSelectExpression.GetDataRowColumnName
        If String.IsNullOrWhiteSpace([Alias]) Then
            Throw New BTSqlException("If using as a select expression, you must have an alias.")
        End If
        Return [Alias]
    End Function

    Public Function RenderForFilter() As String Implements ISqlSelectQuery.RenderForFilter, ISqlSelectExpression.RenderForFilter, IBTSqlSelectQuery.RenderForFilter
        Dim sb As New StringBuilder()
        sb.Append("(")
        sb.Append(_sqlBuilder.RenderForFilter())
        sb.Append(")")
        Return sb.ToString()
    End Function

    Public Function RenderForGroupBy() As String Implements ISqlSelectExpression.RenderForGroupBy
        Return GetDataRowColumnName()
    End Function

    Public Function RenderForOrderBy() As String Implements ISqlSelectExpression.RenderForOrderBy
        Return GetDataRowColumnName()
    End Function

    Public Function RenderForFunction() As String Implements ISqlSelectExpression.RenderForFunction
        Return RenderForAssignment()
    End Function

    Public Function RenderForJoin() As String Implements ISqlSelectExpression.RenderForJoin
        Throw New BTSqlException("Select Query cannot be used in a join clause.")
    End Function

    Public Function GetParameters() As List(Of IBTSqlParameter) Implements ISqlSelectExpression.GetParameters
        Return _sqlBuilder.GetParameters()
    End Function

    Public ReadOnly Property Parameters As IEnumerable(Of SqlParameter) Implements IHasParameters.Parameters
        Get
            Return _sqlBuilder.Parameters
        End Get
    End Property

    Public Function Render() As String Implements ISqlSelectQuery.Render
        Return String.Format("({0}) AS {1}", _sqlBuilder.Render(), GetDataRowColumnName)
    End Function

    Public Function RenderForAssignment() As String Implements ISqlSelectQuery.RenderForAssignment
        Return RenderForFilter()
    End Function

    Public Function GetDependencyIdentifiers() As List(Of String) Implements ISqlSelectQuery.GetDependencyIdentifiers
        Return _sqlBuilder.GetDependencyIdentifiers()
    End Function

    Public Function GetDependenciesByIdentifier(ByVal owner As String) As List(Of ISqlSelectExpression) Implements ISqlSelectQuery.GetDependenciesByIdentifier
        Return _sqlBuilder.GetDependenciesByIdentifier(owner)
    End Function

    Public Function RenderForUnion() As String Implements ISqlUnionable.RenderForUnion
        Return _sqlBuilder.Render()
    End Function

    Public Function Clone() As Object Implements ICloneable.Clone
        'NOTE: this performs a shallow copy
        Return New BTSqlSelectQuery(_sqlBuilder)
    End Function

    Private _isTableParameterTable As Boolean = False
    Public ReadOnly Property IsTableParameterTable As Boolean Implements ISqlSelectQuery.IsTableParameterTable
        Get
            Return _isTableParameterTable
        End Get
    End Property

End Class

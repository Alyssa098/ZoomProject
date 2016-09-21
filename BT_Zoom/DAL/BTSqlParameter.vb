Option Explicit On
Option Strict On

Imports System.Collections.Generic
Imports System.Data.SqlClient
Imports System.Text
Imports BT_Zoom.Interfaces


Public Class BTSqlParameter
    Implements IBTSqlParameter

    Public Sub New(ByVal parameter As SqlParameter, Optional ByVal [alias] As String = "")
        Me._Parameter = parameter
        Me.Alias = [alias]
    End Sub

    Public Sub New(paramName As String, value As Object, Optional [alias] As String = "")
        Me._Parameter = New SqlParameter(paramName, value)
        Me.Alias = [alias]
    End Sub

    Public Property Parameter As SqlParameter Implements IBTSqlParameter.Parameter

    Public Property [Alias] As String Implements ISqlSelectExpression.Alias

    Public Function WithAlias([alias] As String) As ISqlSelectExpression Implements ISqlSelectExpression.WithAlias
        Dim result As BTSqlParameter = CType(Clone(), BTSqlParameter)
        result.Alias = [alias]
        Return result
    End Function

    Public Function GetDataRowColumnName() As String Implements ISqlSelectExpression.GetDataRowColumnName
        If Not String.IsNullOrWhitespace(Me.Alias) Then
            Return Me.Alias.Replace("[", String.Empty).Replace("]", String.Empty)
        Else
            Return Parameter.ParameterName
        End If
    End Function

#Region "ToString / Render"

    Public Overrides Function ToString() As String Implements IBTSqlParameter.ToString
        Return Parameter.ParameterName
    End Function

    Public Function Render() As String Implements ISqlExpression.Render
        Return String.Format("{0}{1}", ToString(), If(Not String.IsNullOrWhitespace(Me.Alias), String.Format(" AS {0}", Me.Alias), String.Empty))
    End Function

    Public Function RenderForFilter() As String Implements ISqlSelectExpression.RenderForFilter
        Return ToString()
    End Function

    Public Function RenderForGroupBy() As String Implements ISqlSelectExpression.RenderForGroupBy
        Return ToString()
    End Function

    Public Function RenderForOrderBy() As String Implements ISqlSelectExpression.RenderForOrderBy
        Return ToString()
    End Function

    Public Function RenderForFunction() As String Implements ISqlSelectExpression.RenderForFunction
        Return ToString()
    End Function

    Public Function RenderForJoin() As String Implements ISqlSelectExpression.RenderForJoin
        Throw New BTSqlException("Parameter expression cannot be rendered for join ON clause.")
    End Function

    Public Overridable Function RenderForAssignment() As String Implements ISqlSelectExpression.RenderForAssignment
        Return ToString()
    End Function

#End Region

    Public Function GetParameters() As List(Of IBTSqlParameter) Implements ISqlSelectExpression.GetParameters
        Return New List(Of IBTSqlParameter) From {Me}
    End Function

    Public ReadOnly Property Parameters As IEnumerable(Of SqlParameter) Implements IHasParameters.Parameters
        Get
            Return New List(Of SqlParameter) From {Me.Parameter}
        End Get
    End Property

    Public Function GetDependencyIdentifiers() As List(Of String) Implements IHasDependencies.GetDependencyIdentifiers
        Dim result As New List(Of String)
        Return result
    End Function

    Public Function GetDependenciesByIdentifier(ByVal owner As String) As List(Of ISqlSelectExpression) Implements IHasDependencies.GetDependenciesByIdentifier
        Dim result As New List(Of ISqlSelectExpression)
        Return result
    End Function

    Public Sub UpdateOwnerForPaging(pagingAlias As String, ownersToChange As Collections.Generic.List(Of String)) Implements ISqlSelectExpression.UpdateOwnerForPaging
        'nothing to do
    End Sub

    Public Function Clone() As Object Implements ICloneable.Clone
        'NOTE: this performs a shallow copy
        Return New BTSqlParameter(Parameter, Me.Alias)
    End Function

End Class

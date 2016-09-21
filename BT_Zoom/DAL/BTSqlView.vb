Option Explicit On
Option Strict On

Imports System.Collections.Generic
Imports BT_Zoom.Interfaces

Public Class BTSqlView
    Implements ISqlViewExpression, IBTSqlView

    Private _viewName As String
    Public ReadOnly Property ViewName As String Implements IBTSqlView.ViewName
        Get
            Return _viewName
        End Get
    End Property

    Private _schema As String
    Public ReadOnly Property Schema As String Implements IBTSqlView.Schema
        Get
            Return _schema
        End Get
    End Property

    Private _alias As String
    Public ReadOnly Property [Alias] As String Implements ISqlJoinable.Alias
        Get
            Return _alias
        End Get
    End Property

    Private _databaseName As String
    Public ReadOnly Property DatabaseName As String Implements IBTSqlView.DatabaseName
        Get
            Return _databaseName
        End Get
    End Property

    Private _allowNoRepresentation As Boolean = False
    Public ReadOnly Property AllowNoRepresentation As Boolean Implements IBTSqlView.AllowNoRepresentation
        Get
            Return _allowNoRepresentation
        End Get
    End Property

    Public Property UseDirtyRead As Boolean = False Implements ISqlBaseTableExpression.UseDirtyRead

    Public Function WithAlias(ByVal [alias] As String) As IBTSqlView Implements IBTSqlView.WithAlias
        Return New BTSqlView(Me.ViewName, Me.Schema, [alias])
    End Function

    Public Sub ChangeAlias(ByVal [alias] As String) Implements ISqlJoinable.ChangeAlias
        If String.IsNullOrWhiteSpace(Me.ViewName) AndAlso String.IsNullOrWhiteSpace(Me.Schema) AndAlso String.IsNullOrWhiteSpace([alias]) AndAlso Not Me.AllowNoRepresentation Then
            Throw New BTSqlException("This action would result in the BTSqlView having no representation")
        End If
        _alias = [alias]
    End Sub

    Public Sub New(ByVal name As String, ByVal schema As String, ByVal [alias] As String, Optional ByVal databaseName As String = "", Optional ByVal allowNoRepresentation As Boolean = False)
        _viewName = name
        _schema = schema
        _alias = [alias]
        _databaseName = databaseName
        _allowNoRepresentation = allowNoRepresentation
        If String.IsNullOrWhiteSpace(Me.ViewName) AndAlso String.IsNullOrWhiteSpace(Me.Schema) AndAlso String.IsNullOrWhiteSpace(Me.Alias) AndAlso Not Me.AllowNoRepresentation Then
            Throw New BTSqlException("This action would result in the BTSqlView having no representation")
        End If
    End Sub

    Public Overrides Function ToString() As String
        Return String.Format("{0}{1}{2}{3}",
                             If(Not String.IsNullOrWhiteSpace(DatabaseName), String.Format("{0}.", DatabaseName), String.Empty),
                             If(Not String.IsNullOrWhiteSpace(Schema), String.Format("{0}.", Schema), String.Empty),
                             If(Not String.IsNullOrWhiteSpace(ViewName), ViewName, String.Empty),
                             If(Not String.IsNullOrWhiteSpace(Me.Alias), String.Format(" {0}", Me.Alias), String.Empty))
    End Function

    Public Function Render() As String Implements ISqlExpression.Render
        Return ToString()
    End Function

    Public Function GetDependencyIdentifiers() As List(Of String) Implements ISqlFromExpression.GetDependencyIdentifiers
        Dim result As New List(Of String)
        'no dependencies
        Return result
    End Function

    Public Function GetDependenciesByIdentifier(ByVal owner As String) As List(Of ISqlSelectExpression) Implements ISqlFromExpression.GetDependenciesByIdentifier
        Return New List(Of ISqlSelectExpression) 'no dependencies
    End Function

    Protected Overridable Function AddCustomExpressionsToStar() As List(Of ISqlSelectExpression) Implements IBTSqlView.AddCustomExpressionsToStar
        Return New List(Of ISqlSelectExpression)()
    End Function

    Protected Overridable Sub RemoveAutoExpressionsFromStar(autoExpressions As List(Of ISqlSelectExpression)) Implements IBTSqlView.RemoveAutoExpressionsFromStar

    End Sub

End Class

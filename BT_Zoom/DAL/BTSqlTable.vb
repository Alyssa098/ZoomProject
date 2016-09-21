Option Explicit On
Option Strict On

Imports System.Collections.Generic
Imports BT_Zoom.Interfaces

<Serializable>
Public MustInherit Class BTSqlTableBase
    Implements IBTSqlTableBase

    Public Sub New(ByVal name As String, ByVal [alias] As String)
        _tableName = name
        _alias = [alias]
    End Sub

    Private _tableName As String
    Public ReadOnly Property TableName As String Implements IBTSqlTableBase.TableName
        Get
            Return _tableName
        End Get
    End Property
    
    Private _alias As String
    Public ReadOnly Property [Alias] As String Implements ISqlJoinable.Alias
        Get
            Return _alias
        End Get
    End Property

    Public MustOverride Function WithAlias([alias] As String) As IBTSqlTableBase Implements IBTSqlTableBase.WithAlias

    Public Overridable Sub ChangeAlias([alias] As String) Implements ISqlJoinable.ChangeAlias
        _alias = [alias]
    End Sub

    Public MustOverride Overrides Function ToString() As String Implements IBTSqlTableBase.ToString

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
   
End Class

<Serializable>
Public Class BTSqlTable
    Inherits BTSqlTableBase
    Implements IBTSqlTable

    Private _schema As String
    Public ReadOnly Property Schema As String Implements IBTSqlTable.Schema
        Get
            Return _schema
        End Get
    End Property
    
    Private _databaseName As String
    Public ReadOnly Property DatabaseName As String Implements IBTSqlTable.DatabaseName
        Get
            Return _databaseName
        End Get
    End Property

    Private _allowNoRepresentation As Boolean = False
    Public ReadOnly Property AllowNoRepresentation As Boolean Implements IBTSqlTable.AllowNoRepresentation
        Get
            Return _allowNoRepresentation
        End Get
    End Property

    Public Property UseDirtyRead As Boolean = False Implements ISqlBaseTableExpression.UseDirtyRead

    Private _indexToUse As String
    ''' <summary>
    ''' Set this property to the name of an index to use as an index hint
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks>
    ''' Reference: http://blog.sqlauthority.com/2009/02/07/sql-server-introduction-to-force-index-query-hints-index-hint/
    ''' </remarks>
    Public Property IndexToUse As String Implements IBTSqlTable.IndexToUse
        Get
            Return _indexToUse
        End Get
        Set(value As String)
            _indexToUse = value
        End Set
    End Property

    Public Overrides Function WithAlias(ByVal [alias] As String) As IBTSqlTableBase
        Return New BTSqlTable(Me.TableName, Me.Schema, [alias])
    End Function

    Public Overrides Sub ChangeAlias(ByVal [alias] As String)
        If String.IsNullOrWhiteSpace(Me.TableName) AndAlso String.IsNullOrWhiteSpace(Me.Schema) AndAlso String.IsNullOrWhiteSpace([alias]) AndAlso Not Me.AllowNoRepresentation Then
            Throw New BTSqlException("This action would result in the BTSqlTable having no representation")
        End If
        MyBase.ChangeAlias([alias])
    End Sub

    Public Sub ChangeDatabase(ByVal databaseName As String) Implements IBTSqlTable.ChangeDatabase
        _databaseName = databaseName
    End Sub

    Public Sub New(ByVal name As String, ByVal schema As String, ByVal [alias] As String, Optional ByVal databaseName As String = "", Optional ByVal allowNoRepresentation As Boolean = False)
        MyBase.New(name, [alias])
        _schema = schema
        _databaseName = databaseName
        _allowNoRepresentation = allowNoRepresentation
        If String.IsNullOrWhiteSpace(Me.TableName) AndAlso String.IsNullOrWhiteSpace(Me.Schema) AndAlso String.IsNullOrWhiteSpace(Me.Alias) AndAlso Not Me.AllowNoRepresentation Then
            Throw New BTSqlException("This action would result in the BTSqlTable having no representation")
        End If
    End Sub

    Public Overrides Function ToString() As String
        Return String.Format("{0}{1}{2}{3}{4}{5}",
                             If(Not String.IsNullOrWhiteSpace(DatabaseName), String.Format("{0}.", DatabaseName), String.Empty),
                             If(Not String.IsNullOrWhiteSpace(Schema), String.Format("{0}.", Schema), String.Empty),
                             If(Not String.IsNullOrWhiteSpace(TableName), TableName, String.Empty),
                             If(Not String.IsNullOrWhiteSpace(Me.Alias), String.Format(" {0}", Me.Alias), String.Empty),
                             If(Not String.IsNullOrWhiteSpace(Me.IndexToUse), String.Format(" WITH( INDEX ({0}) )", Me.IndexToUse), String.Empty),
                             If(Me.UseDirtyRead, " WITH(NOLOCK)", String.Empty))
    End Function

    Protected Overridable Function AddCustomExpressionsToStar() As List(Of ISqlSelectExpression) Implements IBTSqlTable.AddCustomExpressionsToStar
        Return New List(Of ISqlSelectExpression)()
    End Function

    Protected Overridable Sub RemoveAutoExpressionsFromStar(autoExpressions As List(Of ISqlSelectExpression)) Implements IBTSqlTable.RemoveAutoExpressionsFromStar

    End Sub

    Protected Overridable Sub AddAliasesWithPrefix(expressions As List(Of ISqlSelectExpression)) Implements IBTSqlTable.AddAliasesWithPrefix

    End Sub

End Class

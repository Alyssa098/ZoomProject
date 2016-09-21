Option Strict On
Option Explicit On

Imports System
Imports System.Collections.Generic
Imports System.Text
Imports System.Data
Imports System.Data.SqlClient
Imports BT_Zoom.Interfaces
Imports BT_Zoom.Interfaces.DAL

Public Class BTAccountingIDsTable
    Inherits BTAccountingIDsTableBase
    Implements IBTAccountingIDsTable

    Public Sub New()
        MyBase.New()
    End Sub

    Public Sub New(ByVal tableAlias As String)
        MyBase.New(tableAlias)
    End Sub

    Public Sub New(ByVal tableAlias As String, columnPrefix As String)
        MyBase.New("AccountingIDs", "dimension", tableAlias)
        _columnPrefix = columnPrefix
    End Sub

    Private _columnPrefix As String = String.Empty
    Public ReadOnly Property ColumnPrefix As String Implements IBTAccountingIDsTable.ColumnPrefix
        Get
            Return _columnPrefix
        End Get
    End Property

    Protected Overrides Sub AddAliasesWithPrefix(expressions As List(Of ISqlSelectExpression))

        If Not String.IsNullOrWhiteSpace(ColumnPrefix) Then
            For Each expression As ISqlSelectExpression In expressions
                expression.Alias = String.Format("{0}{1}", ColumnPrefix, expression.GetDataRowColumnName())
            Next
        End If

    End Sub

End Class

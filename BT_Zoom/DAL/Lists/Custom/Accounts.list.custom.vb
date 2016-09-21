Option Strict On
Option Explicit On

Imports System
Imports System.Collections.Generic
Imports System.Text
Imports System.Data
Imports System.Data.SqlClient
Imports BT_Zoom.Interfaces.DAL

Public Class BTAccountsList
    Inherits BTAccountsListBase(Of BTAccounts, BTAccountsTable, BTAccountsList)

#Region "Overrides"

    Protected Overrides Function CreateEntity() As BTAccounts
        Return New BTAccounts()
    End Function

    Protected Overrides Function CreateTable() As BTAccountsTable
        Return New BTAccountsTable()
    End Function

#End Region

    Public Sub New()
        MyBase.New()
    End Sub

End Class

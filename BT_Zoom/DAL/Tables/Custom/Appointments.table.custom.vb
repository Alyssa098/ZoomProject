Option Strict On
Option Explicit On

Imports System
Imports System.Collections.Generic
Imports System.Text
Imports System.Data
Imports System.Data.SqlClient
Imports BT_Zoom.Interfaces
Imports BT_Zoom.Interfaces.DAL

Public Class BTAppointmentsTable
    Inherits BTAppointmentsTableBase

    Public Sub New()
        MyBase.New()
    End Sub

    Public Sub New(ByVal tableAlias As String)
        MyBase.New(tableAlias)
    End Sub

End Class

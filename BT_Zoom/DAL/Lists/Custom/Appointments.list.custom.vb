Option Strict On
Option Explicit On

Imports System
Imports System.Collections.Generic
Imports System.Text
Imports System.Data
Imports System.Data.SqlClient
Imports BT_Zoom.Interfaces.DAL

Public Class BTAppointmentsList
    Inherits BTAppointmentsListBase(Of BTAppointments, BTAppointmentsTable, BTAppointmentsList)

#Region "Overrides"

    Protected Overrides Function CreateEntity() As BTAppointments
        Return New BTAppointments()
    End Function

    Protected Overrides Function CreateTable() As BTAppointmentsTable
        Return New BTAppointmentsTable()
    End Function

#End Region

    Public Sub New()
        MyBase.New()
    End Sub

End Class

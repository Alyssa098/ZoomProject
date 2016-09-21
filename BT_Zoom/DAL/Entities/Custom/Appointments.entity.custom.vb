Imports System
Imports System.Collections.Generic
Imports System.Text
Imports System.Data
Imports System.Data.SqlClient
Imports System.Runtime.Serialization
Imports BT_Zoom.Enums.Entities
Imports BT_Zoom.Interfaces
Imports BT_Zoom.Interfaces.DAL

Public Class BTAppointments
    Inherits BTAppointmentsBase(Of BTAppointmentsTable, BTAppointmentsList)

#Region "Overrides"

    Protected Overrides Function CreateList() As BTAppointmentsList
        Return New BTAppointmentsList()
    End Function

    Protected Overrides Function CreateTable() As BTAppointmentsTable
        Return New BTAppointmentsTable()
    End Function

    Public Overrides Sub SetAddedBy(userId As Guid, timestamp As Date)
        AddedByID.Obj = userId
        AddedByDate.Utc = timestamp
    End Sub

    Protected Overrides Sub SetUpdatedByInternal(userId As Guid, timestamp As Date)
        UpdatedByID.Obj = userId
        UpdatedByDate.Utc = timestamp
    End Sub

#End Region

#Region "Constructors"

    Public Sub New()
        MyBase.New()
    End Sub

    Public Sub New(ByVal setDefaults As Boolean)
        MyBase.New()
        If setDefaults Then
            SetEntityDefaults()
        End If
    End Sub

    Public Sub New(processID As Integer)
        MyBase.New()
        Initialize(processID)
    End Sub

    Public Sub New(dr As DataRow, ParamArray mappings() As IBTSqlColumnBase)
        MyBase.New()
        PopulateFromDataRow(dr, mappings)
    End Sub

    Public Sub New(ByVal info As SerializationInfo, ByVal context As StreamingContext)
        MyBase.New()
        PopulateFromSerializationInfo(info)
    End Sub

#End Region

End Class

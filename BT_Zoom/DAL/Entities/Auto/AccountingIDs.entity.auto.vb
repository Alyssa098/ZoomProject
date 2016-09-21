'
' This file was automatically generated and will be again.  
' Do not make edits in this file or they will be lost!!!
' Instead, make custom edits to the partial class of the same name with a ".custom.vb" extension.
'
Imports System
Imports System.Collections.Generic
Imports System.Text
Imports System.Data
Imports System.Data.SqlClient
Imports BT_Zoom.Enums.Entities
Imports BT_Zoom.Enums.BTSql
Imports BT_Zoom.Enums
Imports BT_Zoom.Interfaces
Imports BT_Zoom.Interfaces.DAL

<Serializable>
Public MustInherit Class BTAccountingIDsBase(Of TTable As {IBTAccountingIDsTableBase, New}, TList As {IBTBaseList(Of TTable)})
    Inherits BTBaseEntity(Of TTable, TList)
    Implements IBTAccountingIDsBase(Of TTable, TList)

#Region "Constructors"

    Protected Sub New()
        MyBase.New()

        RegisterFields(_fiAccountingID, _fiBuilderID, _fiIntegrationType, _fiIdType, _fiQbV2ID, _fiQbV2IDType, _fiQbV3ID, _fiXeroID, _fiQbSyncStatus, _fiQbSyncStatusLastUpdatedDate)

    End Sub

    Protected Sub Initialize(accountingID As Integer)

        _list = CreateList()
        _list.AddFilter(_list.From.AccountingID, ComparisonOperatorTypes.Equals, _list.AddParameter("@accountingID", accountingID))

        _list.AddSelect(_list.From.Star())
        LoadInternal()

    End Sub

#End Region

#Region "BT Properties"

    Protected _fiAccountingID As ISqlFieldInfo = New SqlFieldInfo() With {.Column = MyBase.Table.AccountingID, .FieldName = "AccountingID", .CreateField = AddressOf CreateBTSqlInt32}
    Public ReadOnly Property AccountingID As IBTSqlInt32 Implements IBTAccountingIDsBase(Of TTable, TList).AccountingID
        Get
            _fiAccountingID.CreateMeIfNecessary()
            Return DirectCast(_fiAccountingID.Field, IBTSqlInt32)
        End Get
    End Property

    Protected _fiBuilderID As ISqlFieldInfo = New SqlFieldInfo() With {.Column = MyBase.Table.BuilderID, .FieldName = "BuilderID", .CreateField = AddressOf CreateBTSqlInt32}
    Public ReadOnly Property BuilderID As IBTSqlInt32 Implements IBTAccountingIDsBase(Of TTable, TList).BuilderID
        Get
            _fiBuilderID.CreateMeIfNecessary()
            Return DirectCast(_fiBuilderID.Field, IBTSqlInt32)
        End Get
    End Property

    Protected _fiIntegrationType As ISqlFieldInfo = New SqlFieldInfo() With {.Column = MyBase.Table.IntegrationType, .FieldName = "IntegrationType", .CreateField = AddressOf CreateBTSqlInt32}
    Public ReadOnly Property IntegrationType As IBTSqlInt32 Implements IBTAccountingIDsBase(Of TTable, TList).IntegrationType
        Get
            _fiIntegrationType.CreateMeIfNecessary()
            Return DirectCast(_fiIntegrationType.Field, IBTSqlInt32)
        End Get
    End Property

    Protected _fiIdType As ISqlFieldInfo = New SqlFieldInfo() With {.Column = MyBase.Table.IdType, .FieldName = "IdType", .CreateField = AddressOf CreateBTSqlInt32}
    Public ReadOnly Property IdType As IBTSqlInt32 Implements IBTAccountingIDsBase(Of TTable, TList).IdType
        Get
            _fiIdType.CreateMeIfNecessary()
            Return DirectCast(_fiIdType.Field, IBTSqlInt32)
        End Get
    End Property

    Protected _fiQbV2ID As ISqlFieldInfo = New SqlFieldInfo() With {.Column = MyBase.Table.QbV2ID, .FieldName = "QbV2ID", .CreateField = AddressOf CreateBTSqlInt64}
    Public ReadOnly Property QbV2ID As IBTSqlInt64 Implements IBTAccountingIDsBase(Of TTable, TList).QbV2ID
        Get
            _fiQbV2ID.CreateMeIfNecessary()
            Return DirectCast(_fiQbV2ID.Field, IBTSqlInt64)
        End Get
    End Property

    Protected _fiQbV2IDType As ISqlFieldInfo = New SqlFieldInfo() With {.Column = MyBase.Table.QbV2IDType, .FieldName = "QbV2IDType", .CreateField = AddressOf CreateBTSqlInt32}
    Public ReadOnly Property QbV2IDType As IBTSqlInt32 Implements IBTAccountingIDsBase(Of TTable, TList).QbV2IDType
        Get
            _fiQbV2IDType.CreateMeIfNecessary()
            Return DirectCast(_fiQbV2IDType.Field, IBTSqlInt32)
        End Get
    End Property

    Protected _fiQbV3ID As ISqlFieldInfo = New SqlFieldInfo() With {.Column = MyBase.Table.QbV3ID, .FieldName = "QbV3ID", .CreateField = AddressOf CreateBTSqlString}
    Public ReadOnly Property QbV3ID As IBTSqlString Implements IBTAccountingIDsBase(Of TTable, TList).QbV3ID
        Get
            _fiQbV3ID.CreateMeIfNecessary()
            Return DirectCast(_fiQbV3ID.Field, IBTSqlString)
        End Get
    End Property

    Protected _fiXeroID As ISqlFieldInfo = New SqlFieldInfo() With {.Column = MyBase.Table.XeroID, .FieldName = "XeroID", .CreateField = AddressOf CreateBTSqlGuid}
    Public ReadOnly Property XeroID As IBTSqlGuid Implements IBTAccountingIDsBase(Of TTable, TList).XeroID
        Get
            _fiXeroID.CreateMeIfNecessary()
            Return DirectCast(_fiXeroID.Field, IBTSqlGuid)
        End Get
    End Property

    Protected _fiQbSyncStatus As ISqlFieldInfo = New SqlFieldInfo() With {.Column = MyBase.Table.QbSyncStatus, .FieldName = "QbSyncStatus", .CreateField = AddressOf CreateBTSqlInt32}
    Public ReadOnly Property QbSyncStatus As IBTSqlInt32 Implements IBTAccountingIDsBase(Of TTable, TList).QbSyncStatus
        Get
            _fiQbSyncStatus.CreateMeIfNecessary()
            Return DirectCast(_fiQbSyncStatus.Field, IBTSqlInt32)
        End Get
    End Property

    Protected _fiQbSyncStatusLastUpdatedDate As ISqlFieldInfo = New SqlFieldInfo() With {.Column = MyBase.Table.QbSyncStatusLastUpdatedDate, .FieldName = "QbSyncStatusLastUpdatedDate", .CreateField = AddressOf CreateBTSqlDateTime}
    Public ReadOnly Property QbSyncStatusLastUpdatedDate As IBTSqlDateTime Implements IBTAccountingIDsBase(Of TTable, TList).QbSyncStatusLastUpdatedDate
        Get
            _fiQbSyncStatusLastUpdatedDate.CreateMeIfNecessary()
            Return DirectCast(_fiQbSyncStatusLastUpdatedDate.Field, IBTSqlDateTime)
        End Get
    End Property

#End Region

End Class

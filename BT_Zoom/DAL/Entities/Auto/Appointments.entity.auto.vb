Imports BT_Zoom.Enums.BTSql
Imports BT_Zoom.Enums.Entities
Imports BT_Zoom.Interfaces
Imports BT_Zoom.Interfaces.DAL

<Serializable>
Public MustInherit Class BTAppointmentsBase(Of TTable As {BTAppointmentsTableBase, New}, TList As {IBTBaseList(Of TTable)})
    Inherits BTBaseEntity(Of TTable, TList)

#Region "Constructors"

    Protected Sub New()
        MyBase.New()

        RegisterFields(_fiAppointmentId, _fiName, _fiAddedByID, _fiAddedByDate, _fiUpdatedByID, _fiUpdatedByDate, _fiAccountId, _fiAppointmentNotes, _fiZoomAppointmentId)

    End Sub

    Protected Sub Initialize(AppointmentId As Integer)

        _list = CreateList()
        _list.AddFilter(_list.From.AppointmentId, ComparisonOperatorTypes.Equals, _list.AddParameter("@AppointmentId", AppointmentId))

        _list.AddSelect(_list.From.Star())
        LoadInternal()

    End Sub

#End Region

#Region "BT Properties"

    Protected _fiAppointmentId As ISqlFieldInfo = New SqlFieldInfo() With {.Column = MyBase.Table.AppointmentId, .FieldName = "AppointmentId", .CreateField = AddressOf CreateBTSqlInt32}
    Public ReadOnly Property AppointmentId As IBTSqlInt32
        Get
            _fiAppointmentId.CreateMeIfNecessary()
            Return DirectCast(_fiAppointmentId.Field, IBTSqlInt32)
        End Get
    End Property

    Protected _fiName As ISqlFieldInfo = New SqlFieldInfo() With {.Column = MyBase.Table.Name, .FieldName = "Name", .CreateField = AddressOf CreateBTSqlString}
    Public ReadOnly Property Name As IBTSqlString
        Get
            _fiName.CreateMeIfNecessary()
            Return DirectCast(_fiName.Field, IBTSqlString)
        End Get
    End Property

    Protected _fiAddedByID As ISqlFieldInfo = New SqlFieldInfo() With {.Column = MyBase.Table.AddedByID, .FieldName = "AddedById", .CreateField = AddressOf CreateBTSqlGuid}
    Public ReadOnly Property AddedByID As IBTSqlGuid
        Get
            _fiAddedByID.CreateMeIfNecessary()
            Return DirectCast(_fiAddedByID.Field, IBTSqlGuid)
        End Get
    End Property

    Protected _fiAddedByDate As ISqlFieldInfo = New SqlFieldInfo() With {.Column = MyBase.Table.AddedByDate, .FieldName = "AddedByDate", .CreateField = AddressOf CreateBTSqlDateTime}
    Public ReadOnly Property AddedByDate As IBTSqlDateTime
        Get
            _fiAddedByDate.CreateMeIfNecessary()
            Return DirectCast(_fiAddedByDate.Field, IBTSqlDateTime)
        End Get
    End Property

    Protected _fiUpdatedByID As ISqlFieldInfo = New SqlFieldInfo() With {.Column = MyBase.Table.UpdatedByID, .FieldName = "UpdatedById", .CreateField = AddressOf CreateBTSqlGuid}
    Public ReadOnly Property UpdatedByID As IBTSqlGuid
        Get
            _fiUpdatedByID.CreateMeIfNecessary()
            Return DirectCast(_fiUpdatedByID.Field, IBTSqlGuid)
        End Get
    End Property

    Protected _fiUpdatedByDate As ISqlFieldInfo = New SqlFieldInfo() With {.Column = MyBase.Table.UpdatedByDate, .FieldName = "UpdatedByDate", .CreateField = AddressOf CreateBTSqlDateTime}
    Public ReadOnly Property UpdatedByDate As IBTSqlDateTime
        Get
            _fiUpdatedByDate.CreateMeIfNecessary()
            Return DirectCast(_fiUpdatedByDate.Field, IBTSqlDateTime)
        End Get
    End Property

    Protected _fiAccountId As ISqlFieldInfo = New SqlFieldInfo() With {.Column = MyBase.Table.AccountId, .FieldName = "AccountId", .CreateField = AddressOf CreateBTSqlInt32}
    Public ReadOnly Property AccountId As IBTSqlInt32
        Get
            _fiAccountId.CreateMeIfNecessary()
            Return DirectCast(_fiAccountId.Field, IBTSqlInt32)
        End Get
    End Property

    Protected _fiAppointmentNotes As ISqlFieldInfo = New SqlFieldInfo() With {.Column = MyBase.Table.AppointmentNotes, .FieldName = "AppointmentNotes", .CreateField = AddressOf CreateBTSqlString}
    Public ReadOnly Property AppointmentNotes As IBTSqlString
        Get
            _fiAppointmentNotes.CreateMeIfNecessary()
            Return DirectCast(_fiAppointmentNotes.Field, IBTSqlString)
        End Get
    End Property

    Protected _fiZoomAppointmentId As ISqlFieldInfo = New SqlFieldInfo() With {.Column = MyBase.Table.ZoomAppointmentId, .FieldName = "ZoomAppointmentId", .CreateField = AddressOf CreateBTSqlInt32}
    Public ReadOnly Property ZoomAppointmentId As IBTSqlInt32
        Get
            _fiZoomAppointmentId.CreateMeIfNecessary()
            Return DirectCast(_fiZoomAppointmentId.Field, IBTSqlInt32)
        End Get
    End Property

#End Region

End Class

Imports BT_Zoom.Enums.BTSql
Imports BT_Zoom.Enums.Entities
Imports BT_Zoom.Interfaces
Imports BT_Zoom.Interfaces.DAL

<Serializable>
Public MustInherit Class BTAccountsBase(Of TTable As {BTAccountsTableBase, New}, TList As {IBTBaseList(Of TTable)})
    Inherits BTBaseEntity(Of TTable, TList)

#Region "Constructors"

    Protected Sub New()
        MyBase.New()

        RegisterFields(_fiAccountId, _fiName, _fiAddedByID, _fiAddedByDate, _fiUpdatedByID, _fiUpdatedByDate)

    End Sub

    Protected Sub Initialize(accountId As Integer)
        _list = CreateList()
        _list.AddFilter(_list.From.AccountId, ComparisonOperatorTypes.Equals, _list.AddParameter("@accountId", accountId))

        _list.AddSelect(_list.From.Star())
        LoadInternal()


    End Sub

#End Region

#Region "BT Properties"

    Protected _fiAccountId As ISqlFieldInfo = New SqlFieldInfo() With {.Column = MyBase.Table.AccountId, .FieldName = "AccountId", .CreateField = AddressOf CreateBTSqlInt32}
    Public ReadOnly Property AccountId As IBTSqlInt32
        Get
            _fiAccountId.CreateMeIfNecessary()
            Return DirectCast(_fiAccountId.Field, IBTSqlInt32)
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

#End Region

End Class

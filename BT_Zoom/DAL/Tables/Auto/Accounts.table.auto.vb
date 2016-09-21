Imports BT_Zoom.Interfaces
Imports BT_Zoom.Interfaces.DAL

<Serializable>
Public MustInherit Class BTAccountsTableBase
    Inherits BTSqlTable

    Protected Sub New()
        Me.New("Accounts", "dbo", "acc")
    End Sub

    Protected Sub New(ByVal tableAlias As String)
        Me.New("Accounts", "dbo", tableAlias)
    End Sub

    Protected Sub New(ByVal name As String, ByVal schema As String, ByVal [Alias] As String, Optional ByVal databaseName As String = "", Optional ByVal allowNoRepresentation As Boolean = False)
        MyBase.New(name, schema, [Alias], databaseName, allowNoRepresentation)
    End Sub

    Private ReadOnly _AddedByDate As IBTSqlColumn = New BTSqlColumn(Me, "AddedByDate", SqlDbType.DateTime, 8, False, 23, 3, False, False)
    Public ReadOnly Property AddedByDate As Interfaces.IBTSqlColumn
        Get
            Return _AddedByDate
        End Get
    End Property

    Private ReadOnly _AddedByID As IBTSqlColumn = New BTSqlColumn(Me, "AddedByID", SqlDbType.UniqueIdentifier, 16, True, 0, 0, False, False)
    Public ReadOnly Property AddedByID As Interfaces.IBTSqlColumn
        Get
            Return _AddedByID
        End Get
    End Property

    Private ReadOnly _Name As IBTSqlColumn = New BTSqlColumn(Me, "Name", SqlDbType.VarChar, 50, False, 0, 0, False, False)
    Public ReadOnly Property Name As Interfaces.IBTSqlColumn
        Get
            Return _Name
        End Get
    End Property

    Private ReadOnly _UpdatedByDate As IBTSqlColumn = New BTSqlColumn(Me, "UpdatedByDate", SqlDbType.DateTime, 8, False, 23, 3, False, False)
    Public ReadOnly Property UpdatedByDate As Interfaces.IBTSqlColumn
        Get
            Return _UpdatedByDate
        End Get
    End Property

    Private ReadOnly _UpdatedByID As IBTSqlColumn = New BTSqlColumn(Me, "UpdatedByID", SqlDbType.UniqueIdentifier, 16, True, 0, 0, False, False)
    Public ReadOnly Property UpdatedByID As Interfaces.IBTSqlColumn
        Get
            Return _UpdatedByID
        End Get
    End Property

    Private ReadOnly _AccountId As IBTSqlColumn = New BTSqlColumn(Me, "accountId", SqlDbType.Int, 4, False, 10, 0, True, True)
    Public ReadOnly Property AccountId As Interfaces.IBTSqlColumn
        Get
            Return _AccountId
        End Get
    End Property



    Public Function Star(ParamArray columnsToExlude() As Interfaces.ISqlSelectExpression) As Interfaces.ISqlSelectExpression()
        Return {AccountId, Name, AddedByID, AddedByDate, UpdatedByID, UpdatedByDate}
    End Function

End Class

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
Imports BT_Zoom.Interfaces
Imports BT_Zoom.Interfaces.DAL

<Serializable>
Public MustInherit Class BTAccountingIDsTableBase
    Inherits BTSqlTable
    Implements IBTAccountingIDsTableBase

    Protected Sub New()
        Me.New("AccountingIDs", "dimension", "a")
    End Sub

    Protected Sub New(ByVal tableAlias As String)
        Me.New("AccountingIDs", "dimension", tableAlias)
    End Sub

    Protected Sub New(ByVal name As String, ByVal schema As String, ByVal [Alias] As String, Optional ByVal databaseName As String = "", Optional ByVal allowNoRepresentation As Boolean = False)
        MyBase.New(name, schema, [Alias], databaseName, allowNoRepresentation)
    End Sub

    Private ReadOnly _AccountingID As IBTSqlColumn = New BTSqlColumn(Me, "accountingID", SqlDbType.Int, 4, False, 10, 0, True, True)
    Public ReadOnly Property AccountingID As IBTSqlColumn Implements IBTAccountingIDsTableBase.AccountingID
        Get
            Return _AccountingID
        End Get
    End Property

    Private ReadOnly _BuilderID As IBTSqlColumn = New BTSqlColumn(Me, "builderID", SqlDbType.Int, 4, False, 10, 0, False, False)
    Public ReadOnly Property BuilderID As IBTSqlColumn Implements IBTAccountingIDsTableBase.BuilderID
        Get
            Return _BuilderID
        End Get
    End Property

    Private ReadOnly _IntegrationType As IBTSqlColumn = New BTSqlColumn(Me, "integrationType", SqlDbType.Int, 4, False, 10, 0, False, False)
    Public ReadOnly Property IntegrationType As IBTSqlColumn Implements IBTAccountingIDsTableBase.IntegrationType
        Get
            Return _IntegrationType
        End Get
    End Property

    Private ReadOnly _IdType As IBTSqlColumn = New BTSqlColumn(Me, "idType", SqlDbType.Int, 4, False, 10, 0, False, False)
    Public ReadOnly Property IdType As IBTSqlColumn Implements IBTAccountingIDsTableBase.IdType
        Get
            Return _IdType
        End Get
    End Property

    Private ReadOnly _QbV2ID As IBTSqlColumn = New BTSqlColumn(Me, "qbV2ID", SqlDbType.BigInt, 8, True, 19, 0, False, False)
    Public ReadOnly Property QbV2ID As IBTSqlColumn Implements IBTAccountingIDsTableBase.QbV2ID
        Get
            Return _QbV2ID
        End Get
    End Property

    Private ReadOnly _QbV2IDType As IBTSqlColumn = New BTSqlColumn(Me, "qbV2IDType", SqlDbType.Int, 4, True, 10, 0, False, False)
    Public ReadOnly Property QbV2IDType As IBTSqlColumn Implements IBTAccountingIDsTableBase.QbV2IDType
        Get
            Return _QbV2IDType
        End Get
    End Property

    Private ReadOnly _QbV3ID As IBTSqlColumn = New BTSqlColumn(Me, "qbV3ID", SqlDbType.VarChar, 22, True, 0, 0, False, False)
    Public ReadOnly Property QbV3ID As IBTSqlColumn Implements IBTAccountingIDsTableBase.QbV3ID
        Get
            Return _QbV3ID
        End Get
    End Property

    Private ReadOnly _XeroID As IBTSqlColumn = New BTSqlColumn(Me, "xeroID", SqlDbType.UniqueIdentifier, 16, True, 0, 0, False, False)
    Public ReadOnly Property XeroID As IBTSqlColumn Implements IBTAccountingIDsTableBase.XeroID
        Get
            Return _XeroID
        End Get
    End Property

    Private ReadOnly _QbSyncStatus As IBTSqlColumn = New BTSqlColumn(Me, "qbSyncStatus", SqlDbType.Int, 4, True, 10, 0, False, False)
    Public ReadOnly Property QbSyncStatus As IBTSqlColumn Implements IBTAccountingIDsTableBase.QbSyncStatus
        Get
            Return _QbSyncStatus
        End Get
    End Property

    Private ReadOnly _QbSyncStatusLastUpdatedDate As IBTSqlColumn = New BTSqlColumn(Me, "qbSyncStatusLastUpdatedDate", SqlDbType.DateTime, 8, True, 23, 3, False, False)
    Public ReadOnly Property QbSyncStatusLastUpdatedDate As IBTSqlColumn Implements IBTAccountingIDsTableBase.QbSyncStatusLastUpdatedDate
        Get
            Return _QbSyncStatusLastUpdatedDate
        End Get
    End Property

    'Primary Keys: { AccountingID }

    Public Function Star(ParamArray columnsToExlude As ISqlSelectExpression()) As ISqlSelectExpression() Implements IBTAccountingIDsTableBase.Star
        Dim result As New List(Of ISqlSelectExpression) From {AccountingID, BuilderID, IntegrationType, IdType, QbV2ID, QbV2IDType, QbV3ID, XeroID, QbSyncStatus, QbSyncStatusLastUpdatedDate}

        RemoveAutoExpressionsFromStar(result) 'NOTE: this provides the opportunity for the custom version of the table to remove any of the auto generated columns from Star

        result.AddRange(AddCustomExpressionsToStar()) 'NOTE: this provides the opportunity for the custom version of the table to add other expressions to Star

        If columnsToExlude IsNot Nothing Then
            Dim excludeList As New List(Of ISqlSelectExpression)(columnsToExlude)
            result.RemoveAll(Function(expr As ISqlSelectExpression) excludeList.Contains(expr))
        End If
        
        AddAliasesWithPrefix(result)

        Return result.ToArray()
    End Function

End Class

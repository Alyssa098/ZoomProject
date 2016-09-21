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

Namespace Interfaces

    Namespace DAL

        Public Interface IBTAccountingIDsTableBase
            Inherits IBTSqlTable

            ReadOnly Property AccountingID As IBTSqlColumn
            ReadOnly Property BuilderID As IBTSqlColumn
            ReadOnly Property IntegrationType As IBTSqlColumn
            ReadOnly Property IdType As IBTSqlColumn
            ReadOnly Property QbV2ID As IBTSqlColumn
            ReadOnly Property QbV2IDType As IBTSqlColumn
            ReadOnly Property QbV3ID As IBTSqlColumn
            ReadOnly Property XeroID As IBTSqlColumn
            ReadOnly Property QbSyncStatus As IBTSqlColumn
            ReadOnly Property QbSyncStatusLastUpdatedDate As IBTSqlColumn

            Function Star(ParamArray columnsToExlude As ISqlSelectExpression()) As ISqlSelectExpression()

        End Interface

        Public Interface IBTAccountingIDsListBase(Of TEntity As {IBTBaseEntity(Of TTable, TList)}, TTable As {IBTSqlTable}, TList As {IBTBaseList(Of TTable)})
            Inherits IBTBaseEntityList(Of TEntity, TTable, TList)

        End Interface

        Public Interface IBTAccountingIDsBase(Of TTable As {IBTAccountingIDsTableBase}, TList As {IBTBaseList(Of TTable)})
            Inherits IBTBaseEntity(Of TTable, TList)

            ReadOnly Property AccountingID As IBTSqlInt32
            ReadOnly Property BuilderID As IBTSqlInt32
            ReadOnly Property IntegrationType As IBTSqlInt32
            ReadOnly Property IdType As IBTSqlInt32
            ReadOnly Property QbV2ID As IBTSqlInt64
            ReadOnly Property QbV2IDType As IBTSqlInt32
            ReadOnly Property QbV3ID As IBTSqlString
            ReadOnly Property XeroID As IBTSqlGuid
            ReadOnly Property QbSyncStatus As IBTSqlInt32
            ReadOnly Property QbSyncStatusLastUpdatedDate As IBTSqlDateTime

        End Interface

    End Namespace

End Namespace

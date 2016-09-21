
Imports System
Imports System.Collections.Generic
Imports System.Text
Imports System.Data
Imports System.Data.SqlClient

Namespace Interfaces

    Namespace DAL

        Public Interface IBTAccountingIDsTable
            Inherits IBTAccountingIDsTableBase

            ReadOnly Property ColumnPrefix As String

        End Interface

        Public Interface IBTAccountingIDsList
            Inherits IBTAccountingIDsListBase(Of IBTAccountingIDs, IBTAccountingIDsTable, IBTAccountingIDsList)

        End Interface

        Public Interface IBTAccountingIDs
            Inherits IBTAccountingIDsBase(Of IBTAccountingIDsTable, IBTAccountingIDsList)
        End Interface

    End Namespace

End Namespace

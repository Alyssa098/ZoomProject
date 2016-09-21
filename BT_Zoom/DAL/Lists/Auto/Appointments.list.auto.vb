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
Public MustInherit Class BTAppointmentsListBase(Of TEntity As {IBTBaseEntity(Of TTable, TList)}, TTable As {IBTSqlTable, New}, TList As {IBTBaseList(Of TTable)})
    Inherits BTBaseEntityList(Of TEntity, TTable, TList)
    Implements IBTAccountingIDsListBase(Of TEntity, TTable, TList)

    Public Sub New()
        MyBase.New()
    End Sub

    Public Sub New(ByVal fromTable As TTable)
        MyBase.New(fromTable)
    End Sub

End Class

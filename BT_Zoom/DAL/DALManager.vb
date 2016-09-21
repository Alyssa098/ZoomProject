Imports BT_Zoom.Constants
Imports BT_Zoom.Interfaces

Public Class DALManager

    Public Shared Sub RemoveInfiniteScrollColumns(dt As DataTable)
        If dt.Columns.Contains(BTConstants.DAL.PagingRowIdColumnName) Then
            dt.Columns.Remove(BTConstants.DAL.PagingRowIdColumnName)
        End If

        If dt.Columns.Contains(BTConstants.DAL.PagingTotalRowsColumnName) Then
            dt.Columns.Remove(BTConstants.DAL.PagingTotalRowsColumnName)
        End If
    End Sub

End Class
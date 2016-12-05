Imports System.Data.SqlClient
Imports System.Data.Sql



Public Class SQLControl
    Public SQLCon As New SqlConnection With {.ConnectionString = "Server=unozoomcapstone.database.windows.net;Database=UnoZoomCapstone;User=capstone;Pwd=UnoMavericks1;"}
    Public SQLcmd As SqlCommand

    Public DBDA As SqlDataAdapter
    Public DBTB As DataTable

    Public Params As New List(Of SqlParameter)

    Public RecordCount As Integer
    Public Exceptions As String

    Public Sub New()
    End Sub


    Public Sub ExecQuery(Query As String)
        RecordCount = 0
        Exceptions = ""

        Try
            SQLCon.Open()
            SQLcmd = New SqlCommand(Query, SQLCon)
            Params.ForEach(Sub(p) SQLcmd.Parameters.Add(p))
            Params.Clear()
            DBTB = New DataTable
            DBDA = New SqlDataAdapter(SQLcmd)
            RecordCount = DBDA.Fill(DBTB)

        Catch ex As Exception
            Exceptions = "ExecQuery Error:" & vbNewLine & ex.Message

        Finally
            If SQLCon.State = ConnectionState.Open Then SQLCon.Close()
        End Try

    End Sub



    Public Sub AddParam(Name As String, value As Object)
        Dim NewParam As New SqlParameter(Name, value)
        Params.Add(NewParam)
    End Sub


    Public Function HasConnection() As Boolean
        Try
            SQLCon.Open()

            SQLCon.Close()
            Return True
        Catch ex As Exception
            MsgBox(ex.Message)
            Return False
        End Try
    End Function

    Public Function HasException(Optional Report As Boolean = False) As Boolean
        If String.IsNullOrEmpty(Exceptions) Then Return False
        If Report = True Then MsgBox(Exceptions, MsgBoxStyle.Critical, "Exception:")
        Return True

    End Function

End Class

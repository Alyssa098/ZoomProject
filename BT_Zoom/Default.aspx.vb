Imports System.IO
Imports System.Net
Public Class _Default
    Inherits Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As EventArgs) Handles Me.Load

    End Sub

    Protected Sub TestButton_Click(sender As Object, e As EventArgs) Handles TestButton.Click

        Dim request As WebRequest = WebRequest.Create(New Uri("https://api.zoom.us/v1/user/getbyemail"))
        Dim key = "za1FM1xWS46_Q-RYf1uWVQ"
        Dim secret As String = "mdEvDof5IOGHDKxP0ZY1o8Sr3vStM11czTE2"
        Dim email As String = "zoomcapstone@buildertrend.com"
        Dim data As String = "api_key=za1FM1xWS46_Q-RYf1uWVQ&api_secret=mdEvDof5IOGHDKxP0ZY1o8Sr3vStM11czTE2&email=zoomcapstone@buildertrend.com&data_type=JSON&login_type=100"
        Dim contentType As String = "application/x-www-form-urlencoded"
        System.Net.ServicePointManager.Expect100Continue = False




        request.Method = "POST"
        request.ContentType = contentType


        Dim requestWriter As StreamWriter = New StreamWriter(request.GetRequestStream())
        requestWriter.Write(data)
        requestWriter.Flush()
        requestWriter.Close()

        Dim wResponse As WebResponse = request.GetResponse()
        Dim sResponse As String = ""

        Using srRead As New StreamReader(wResponse.GetResponseStream())
            sResponse = srRead.ReadToEnd()

            Response.Write(sResponse)
        End Using



    End Sub
End Class
Imports System.IO
Imports System.Net
Public Class _Default
    Inherits Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As EventArgs) Handles Me.Load

    End Sub

    Protected Sub TestButton_Click(sender As Object, e As EventArgs) Handles TestButton.Click

        If Not IsNumeric(HourTextBox.Text) Or Not IsNumeric(MinTextBox.Text) Or Not IsNumeric(SecTextBox.Text) Then


        End If

        Dim request As WebRequest = WebRequest.Create(New Uri("https://api.zoom.us/v1/meeting/create"))
        Dim key = "za1FM1xWS46_Q-RYf1uWVQ"
        Dim secret As String = "mdEvDof5IOGHDKxP0ZY1o8Sr3vStM11czTE2"
        Dim hostid As String = "NXspN9OLS16senKZcxDF-g"
        Dim email As String = "zoomcapstone@buildertrend.com"
        'Dim data As String = "api_key=tBe1Vq-NT62DcJVmyt76Ng&api_secret=XMe8Whs0gViqzchgdXrLOMXETOFganhoJ5Nl&host_id=NXspN9OLS16senKZcxDF-g&topic=test&type=2&start_time=2016-11-20T12:00:00Z"
        Dim contentType As String = "application/x-www-form-urlencoded"
        Dim MeetingT As New Date(Convert.ToInt32(MeetingCalendar.SelectedDate.Year), Convert.ToInt32(MeetingCalendar.SelectedDate.Month), Convert.ToInt32(MeetingCalendar.SelectedDate.Day), Convert.ToInt32(HourTextBox.Text), Convert.ToInt32(MinTextBox.Text), Convert.ToInt32(SecTextBox.Text))
        System.Net.ServicePointManager.Expect100Continue = False

        MeetingT = MeetingT.ToUniversalTime()
        Dim data As String = "api_key=tBe1Vq-NT62DcJVmyt76Ng&api_secret=XMe8Whs0gViqzchgdXrLOMXETOFganhoJ5Nl&host_id=NXspN9OLS16senKZcxDF-g" + "&" + "topic=" + TopicTextBox.Text + "&type=2&start_time=" + CStr(MeetingT.Year) + "-" + CStr(MeetingT.Month) + "-" + CStr(MeetingT.Day) + "T" + CStr(MeetingT.Hour) + ":" + CStr(MeetingT.Minute) + ":" + CStr(MeetingT.Second) + "Z"

        MeetingTimeTextBox.Text = CStr(MeetingT.Year) + "-" + CStr(MeetingT.Month) + "-" + CStr(MeetingT.Day) + "      " + CStr(MeetingT.Hour) + ":" + CStr(MeetingT.Minute) + ":" + CStr(MeetingT.Second)
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

    Protected Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim request As WebRequest = WebRequest.Create(New Uri("https://api.zoom.us/v1/meeting/list"))
        Dim key = "za1FM1xWS46_Q-RYf1uWVQ"
        Dim secret As String = "mdEvDof5IOGHDKxP0ZY1o8Sr3vStM11czTE2"
        Dim hostid As String = "NXspN9OLS16senKZcxDF-g"
        Dim email As String = "zoomcapstone@buildertrend.com"
        'Dim data As String = "api_key=tBe1Vq-NT62DcJVmyt76Ng&api_secret=XMe8Whs0gViqzchgdXrLOMXETOFganhoJ5Nl&host_id=NXspN9OLS16senKZcxDF-g&topic=test&type=2&start_time=2016-11-20T12:00:00Z"
        Dim contentType As String = "application/x-www-form-urlencoded"
        System.Net.ServicePointManager.Expect100Continue = False

        Dim data As String = "api_key=tBe1Vq-NT62DcJVmyt76Ng&api_secret=XMe8Whs0gViqzchgdXrLOMXETOFganhoJ5Nl&host_id=NXspN9OLS16senKZcxDF-g"

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

    Protected Sub DropDownList1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles DropDownList1.SelectedIndexChanged

    End Sub
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim request As WebRequest = WebRequest.Create(New Uri("https://api.zoom.us/v1/meeting/list"))
        Dim key = "za1FM1xWS46_Q-RYf1uWVQ"
        Dim secret As String = "mdEvDof5IOGHDKxP0ZY1o8Sr3vStM11czTE2"
        Dim hostid As String = "NXspN9OLS16senKZcxDF-g"
        Dim email As String = "zoomcapstone@buildertrend.com"
        'Dim data As String = "api_key=tBe1Vq-NT62DcJVmyt76Ng&api_secret=XMe8Whs0gViqzchgdXrLOMXETOFganhoJ5Nl&host_id=NXspN9OLS16senKZcxDF-g&topic=test&type=2&start_time=2016-11-20T12:00:00Z"
        Dim contentType As String = "application/x-www-form-urlencoded"
        System.Net.ServicePointManager.Expect100Continue = False

        Dim data As String = "api_key=tBe1Vq-NT62DcJVmyt76Ng&api_secret=XMe8Whs0gViqzchgdXrLOMXETOFganhoJ5Nl&host_id=NXspN9OLS16senKZcxDF-g"

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
            DropDownList1.Items.Add(sResponse)
            Response.Write(sResponse)
        End Using
    End Sub
End Class
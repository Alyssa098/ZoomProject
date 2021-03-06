﻿Imports System.Data.SqlClient
Imports System.Data.Sql
Imports System.IO
Imports System.Net
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports System.Activities.Expressions

Public Class Update
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

    End Sub

    Protected Sub UpdateButton_Click(sender As Object, e As EventArgs) Handles UpdateButton.Click
        Dim ZoomId As String
        ZoomId = CType(Session.Item("zoomId"), String)


        Dim request As WebRequest = WebRequest.Create(New Uri("https://api.zoom.us/v1/meeting/update"))
        Dim contentType As String = "application/x-www-form-urlencoded"
        Dim MeetingT As New Date(Convert.ToInt32(MeetingCalendar.SelectedDate.Year), Convert.ToInt32(MeetingCalendar.SelectedDate.Month), Convert.ToInt32(MeetingCalendar.SelectedDate.Day), Convert.ToInt32(HourTextBox.Text), Convert.ToInt32(MinTextBox.Text), Convert.ToInt32(SecTextBox.Text))
        MeetingT = MeetingT.ToUniversalTime()
        Dim data As String = "api_key=tBe1Vq-NT62DcJVmyt76Ng&api_secret=XMe8Whs0gViqzchgdXrLOMXETOFganhoJ5Nl&id=" + ZoomId + "&host_id=NXspN9OLS16senKZcxDF-g" + "&" + "topic=" + TopicTextBox.Text + "&type=2&start_time=" + CStr(MeetingT.Year) + "-" + CStr(MeetingT.Month) + "-" + CStr(MeetingT.Day) + "T" + CStr(MeetingT.Hour) + ":" + CStr(MeetingT.Minute) + ":" + CStr(MeetingT.Second) + "Z"

        request.Method = "POST"
        request.ContentType = contentType

        Dim requestWriter As StreamWriter = New StreamWriter(request.GetRequestStream())
        requestWriter.Write(data)
        requestWriter.Flush()
        requestWriter.Close()

        Dim wResponse As WebResponse = request.GetResponse()

        Dim qtime As String = CStr(MeetingCalendar.SelectedDate.Year) + "-" + CStr(MeetingCalendar.SelectedDate.Month) + "-" + CStr(MeetingCalendar.SelectedDate.Day) + " " + HourTextBox.Text + ":" + MinTextBox.Text + ":" + SecTextBox.Text
        Dim SQL As New SQLControl
        SQL.AddParam("@Time", qtime)
        SQL.AddParam("@Topic", TopicTextBox.Text)
        SQL.AddParam("@zoomId", ZoomId)
        SQL.ExecQuery("update appointments set AppointmentDate = @Time, AppointmentNotes = @Topic, UpdatedByDate = GETDATE() where ZoomAppointmentId = @zoomId")

        MsgBox("Update Successed!")
        Server.Transfer("Default.aspx", True)
    End Sub

    Protected Sub DeleteButton_Click(sender As Object, e As EventArgs) Handles DeleteButton.Click
        Dim ZoomId As String
        ZoomId = CType(Session.Item("zoomId"), String)

        Dim request As WebRequest = WebRequest.Create(New Uri("https://api.zoom.us/v1/meeting/delete"))
        Dim contentType As String = "application/x-www-form-urlencoded"

        Dim data As String = "api_key=tBe1Vq-NT62DcJVmyt76Ng&api_secret=XMe8Whs0gViqzchgdXrLOMXETOFganhoJ5Nl&id=" + ZoomId + "&host_id=NXspN9OLS16senKZcxDF-g"

        request.Method = "POST"
        request.ContentType = contentType

        Dim requestWriter As StreamWriter = New StreamWriter(request.GetRequestStream())
        requestWriter.Write(data)
        requestWriter.Flush()
        requestWriter.Close()
        Dim wResponse As WebResponse = request.GetResponse()


        Dim SQL As New SQLControl
        SQL.AddParam("@zoomId", ZoomId)
        SQL.ExecQuery("delete from appointments where ZoomAppointmentId = @zoomId")
        MsgBox("Meeting has been canclled!")
        Server.Transfer("Default.aspx", True)
    End Sub

    Protected Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim uuid As String
        uuid = CType(Session.Item("uuid"), String)
        Dim ZoomId As String
        ZoomId = CType(Session.Item("zoomId"), String)

        Dim request As WebRequest = WebRequest.Create(New Uri("https://api.zoom.us/v1/recording/list"))
        Dim contentType As String = "application/x-www-form-urlencoded"

        Dim data As String = "api_key=tBe1Vq-NT62DcJVmyt76Ng&api_secret=XMe8Whs0gViqzchgdXrLOMXETOFganhoJ5Nl&host_id=NXspN9OLS16senKZcxDF-g"  'meeting_id=" + uuid
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
            'Response.Write(sResponse)
        End Using
        Dim ser As JObject = JObject.Parse(sResponse)
        Dim info As List(Of JToken) = ser.Children().ToList
        Dim url As String = ""
        For Each item As JProperty In info
            item.CreateReader()
            Select Case item.Name
                Case "meetings"
                    For Each meetings As JObject In item.Values

                        Dim meetingId As String = meetings("meeting_number")

                        If String.Compare(meetingId, ZoomId) = 0 Then
                            Dim things As List(Of JToken) = meetings.Children().ToList
                            For Each pro As JProperty In things
                                pro.CreateReader()
                                Select Case pro.Name
                                    Case "recording_files"



                                        For Each a As JObject In pro.Values
                                            url = a("play_url")
                                            Exit For
                                        Next
                                End Select
                            Next




                        End If
                    Next
            End Select
        Next


        Response.Redirect(url)


    End Sub
End Class
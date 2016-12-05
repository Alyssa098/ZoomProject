Imports System.Data.SqlClient
Imports System.Data.Sql
Imports System.IO
Imports System.Net
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports System.Activities.Expressions

Public Class _Default
    Inherits Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As EventArgs) Handles Me.Load

    End Sub

    Protected Sub TestButton_Click(sender As Object, e As EventArgs) Handles TestButton.Click


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

        Dim ser As JObject = JObject.Parse(sResponse)

        Dim ZoomAppointmentId As String = ser("id")

        Dim SQL As New SQLControl

        Dim Name As String = NameBox.Text.ToString
        Dim query As String = "select accountId from accounts where name = " + "'" + Name + "'"
        SQL.ExecQuery(query)
        Dim accountId As String
        Dim qnote As String = TopicTextBox.Text

        For Each r As DataRow In SQL.DBTB.Rows
            accountId = r("accountId")
        Next
        Dim qtime As String = CStr(MeetingCalendar.SelectedDate.Year) + "-" + CStr(MeetingCalendar.SelectedDate.Month) + "-" + CStr(MeetingCalendar.SelectedDate.Day) + " " + HourTextBox.Text + ":" + MinTextBox.Text + ":" + SecTextBox.Text

        ' MsgBox(ZoomAppointmentId)
        SQL.AddParam("@Name", Name)
        SQL.AddParam("@AccountID", accountId)
        SQL.AddParam("@AppointmentDate", qtime)
        SQL.AddParam("@AppointmentNotes", qnote)
        SQL.AddParam("@ZoomAppointmentId", ZoomAppointmentId)
        SQL.ExecQuery("insert into Appointments ( Name, AccountID, AppointmentDate, AppointmentNotes, ZoomAppointmentId, AddedByDate, UpdatedByDate)" &
            "VALUES (@name,@AccountID , @AppointmentDate, @AppointmentNotes, @ZoomAppointmentId,GETDATE(), GETDATE());")

        'SQL.ExecQuery(query)





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
        Dim ser As JObject = JObject.Parse(sResponse)
        Dim info As List(Of JToken) = ser.Children().ToList
        Dim output As String = ""

        For Each item As JProperty In info
            item.CreateReader()
            Select Case item.Name
                Case "meetings"
                    For Each meetings As JObject In item.Values
                        Dim ti As String = meetings("start_time")
                        Dim topic As String = meetings("topic")
                        Dim meetingId As String = meetings("id")
                        output += topic + vbTab + ti + vbTab + meetingId + vbCrLf

                    Next
            End Select
        Next


        Dim SQL As New SQLControl
        SQL.ExecQuery("select * from Accounts")


    End Sub

    Protected Sub DropDownList1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles DropDownList1.SelectedIndexChanged
        Dim SQL As New SQLControl
        Dim SelectedName As String = DropDownList1.SelectedValue.ToString

    End Sub
End Class
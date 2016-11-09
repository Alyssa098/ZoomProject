Public Class Login
    Inherits Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As EventArgs) Handles Me.Load
        RegisterHyperLink.NavigateUrl = "Register"

        Dim returnUrl = HttpUtility.UrlEncode(Request.QueryString("ReturnUrl"))
        If Not String.IsNullOrEmpty(returnUrl) Then
            RegisterHyperLink.NavigateUrl &= "?ReturnUrl=" & returnUrl
        End If
    End Sub

    Protected Sub Unnamed6_Click(sender As Object, e As EventArgs)

    End Sub
End Class
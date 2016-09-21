Imports System.Web.Optimization

Public Class Global_asax
    Inherits HttpApplication

    Sub Application_Start(ByVal sender As Object, ByVal e As EventArgs)
        ' Fires when the application is started
        BundleConfig.RegisterBundles(BundleTable.Bundles)
        AuthConfig.RegisterOpenAuth()
        RouteConfig.RegisterRoutes(RouteTable.Routes)

        '''' Seed data. Should really only run once
        Using al As New BTAccountsList
            ' Just checking to see if any accounts exist in the DB. If there are some, we won't do anything. If there are none, we will generate about 20 accounts to use.
            al.AllowNoFilters = True
            al.AddSelect(al.From.Star)
            If al.LoadTop(1) = 0 Then
                Dim accountNames As String() = {"Pancake Man", "Nebraska Corn And Feed", "A-Dell", "Paper Cups Unl.", "Dunder Mifflin", "Michael Scott Paper Company", "Schrute Farms", "Napster", "MySpace", "Enron",
                    "Northwind", "Intamd", "Innotech", "Something With Pirates", "4 Narnia", "The Shire", "BuilderTron"}
                For Each acc As String In accountNames
                    Dim account As New BTAccounts
                    account.Name.Obj = acc ' Name of the account
                    account.SetAddedBy(New Guid, DateTime.UtcNow) ' Just use a blank Guid. Not super important for this project
                    account.SetUpdatedBy(New Guid) ' Just use a blank Guid. Not super important for this project
                    account.Save() ' Create the video. This will also perform an update if the record already exists
                Next
            End If
        End Using

    End Sub

    Sub Application_BeginRequest(ByVal sender As Object, ByVal e As EventArgs)
        ' Fires at the beginning of each request
    End Sub

    Sub Application_AuthenticateRequest(ByVal sender As Object, ByVal e As EventArgs)
        ' Fires upon attempting to authenticate the use
    End Sub

    Sub Application_Error(ByVal sender As Object, ByVal e As EventArgs)
        ' Fires when an error occurs
    End Sub

    Sub Application_End(ByVal sender As Object, ByVal e As EventArgs)
        ' Fires when the application ends
    End Sub
End Class
Imports System.Data.SqlTypes

Public Class DalSample
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        ' This is for adding an Account. You create a new BTAccounts Object, set the name, the added/updated by information and any other information.
        Dim account As New BTAccounts
        account.Name.Obj = "Home Construction Sample" ' Name of the account
        account.SetAddedBy(CType(Membership.GetUser().ProviderUserKey, Guid), DateTime.UtcNow)
        account.SetUpdatedBy(CType(Membership.GetUser().ProviderUserKey, Guid))
        account.Save() ' Create the video. This will also perform an update if the record already exists

        'Query for list of accounts
        Dim accountsList As New BTAccountsList()
        'videosList.AddSelect(videosList.From.Star) ' Will select all fields if you need them all
        accountsList.AddSelect(accountsList.From.Name, accountsList.From.AccountId) ' OR you can do this and it Will only select the fields you specify... 

        'You can choose to add a filter
        accountsList.AddFilter(accountsList.From.Name, Enums.BTSql.ComparisonOperatorTypes.Equals, accountsList.AddParameter("@name", "Your AccountNameHere"))
        accountsList.LoadAll() ' Load all the records

        ' You can enumerate either over all of the rows
        For Each r As DataRow In accountsList.Data.Rows

        Next

        ' Or over strongly-typed representations of the rows (This should be your preferred method).
        For Each a In accountsList.ToList()
            Response.Write(String.Format("{0} <br/>", a.Name.Value))
        Next


        Dim accountOld As New BTAccounts(1) ' If you pass through a valid ID to the constructor, it will load up that instance of the Video object
        ' You can then modify the account
        accountOld.Name.Obj = "New Account Name"
        accountOld.Save()
        'Or you can delete the video (Which I won't do so that the ID 1 is still valid).
        'videoOld.Delete()

    End Sub

End Class
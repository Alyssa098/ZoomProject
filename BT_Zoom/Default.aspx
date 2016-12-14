<%@ Page Title="Home Page" SmartNavigation="true" Language="VB" MasterPageFile="~/Site.Master" AutoEventWireup="true" CodeBehind="Default.aspx.vb" Inherits="BT_Zoom._Default" %>

<asp:Content runat="server" ID="FeaturedContent" ContentPlaceHolderID="FeaturedContent">
    <section class="featured">
        <div class="content-wrapper">
            <hgroup class="title">
                <h1><%: Title %></h1>
                <h2></h2>
            </hgroup>
          
        </div>
    </section>
</asp:Content>
<asp:Content runat="server" ID="BodyContent" ContentPlaceHolderID="MainContent">
   
       
        
    
    <ol>
    <li>
    
        <h4>
            <asp:Label ID="Label1" runat="server" Text="Client Name"></asp:Label>
            <asp:DropDownList ID="NameDropDownList" runat="server" DataSourceID="SqlDataSource1" DataTextField="Name" DataValueField="Name">
            </asp:DropDownList>
        </h4>
    <h4>Topic:<asp:TextBox ID="TopicTextBox" runat="server" Height="16px" Width="347px"></asp:TextBox>
        <asp:RequiredFieldValidator ID="RequiredFieldValidator2" runat="server" ControlToValidate="TopicTextBox" ErrorMessage="Topic is required" ForeColor="Red"></asp:RequiredFieldValidator>
    </h4>
    <h4>Meeing Date:<asp:Calendar ID="MeetingCalendar" runat="server"></asp:Calendar>
    
       <h4>Meeting Time</h4> 

    <asp:TextBox ID="HourTextBox" runat="server" Width="16px" ValidationGroup ="hour"></asp:TextBox>
        :<asp:TextBox ID="MinTextBox" runat="server" Width="16px" ValidationGroup ="time"></asp:TextBox>
        :<asp:TextBox ID="SecTextBox" runat="server" Width="16px" ValidationGroup ="time"></asp:TextBox>
    (24 hour format)
        <asp:Button ID="TestButton" runat="server" Text="Create Meeting" />
        <p>
            &nbsp;</p>
    <h4>&nbsp;</h4>
    </li>
    </ol>
    <ol><li>

        <asp:DropDownList ID="DropDownList1" runat="server"  DataSourceID="SqlDataSource1" DataTextField="Name" DataValueField="Name" Width ="300" AutoPostBack="True">
        </asp:DropDownList>
        <asp:DropDownList ID="DropDownList2" runat="server"  DataSourceID ="SqlDataSource2" DataTextField="AppointmentDate" DataValueField="AppointmentDate" Width ="400" AutoPostBack="True">
        </asp:DropDownList>
        <asp:SqlDataSource ID="SqlDataSource2" runat="server" ConnectionString="Data Source=tcp:unozoomcapstone.database.windows.net,1433;Initial Catalog=UnoZoomCapstone;Persist Security Info=False;User ID=capstone;Password=UnoMavericks1;MultipleActiveResultSets=False;Connect Timeout=30;Encrypt=True;TrustServerCertificate=False" ProviderName="System.Data.SqlClient" SelectCommand="SELECT [AppointmentDate] FROM [Appointments] WHERE ([Name] = @Name)">
            <SelectParameters>
                <asp:ControlParameter ControlID="DropDownList1" Name="Name" PropertyName="SelectedValue" Type="String" />
            </SelectParameters>
        </asp:SqlDataSource>
        <asp:Button ID="Button2" runat="server" Text="Srart Meeting" CausesValidation="false"/>
        <asp:Button ID="Button1" runat="server" Text="Meeting Detail" CausesValidation="false"/>
        <asp:SqlDataSource ID="SqlDataSource1" runat="server" ConnectionString="<%$ ConnectionStrings:UnoZoomCapstoneConnectionString %>" SelectCommand="SELECT [Name] FROM [Accounts]"></asp:SqlDataSource>
        </li></ol>
      </ol>

</asp:Content>
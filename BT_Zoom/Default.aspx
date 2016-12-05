<%@ Page Title="Home Page" Language="VB" MasterPageFile="~/Site.Master" AutoEventWireup="true" CodeBehind="Default.aspx.vb" Inherits="BT_Zoom._Default" %>

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
            <asp:TextBox ID="NameBox" runat="server"></asp:TextBox>
        </h4>
    <h4>Topic:<asp:TextBox ID="TopicTextBox" runat="server" Height="16px" Width="347px"></asp:TextBox>
        <asp:RequiredFieldValidator ID="RequiredFieldValidator2" runat="server" ControlToValidate="TopicTextBox" ErrorMessage="Topic is required" ForeColor="Red"></asp:RequiredFieldValidator>
    </h4>
    <h4>Meeing Date:<asp:Calendar ID="MeetingCalendar" runat="server"></asp:Calendar>
    
       <h4>Meeting Time</h4> 

    <asp:TextBox ID="HourTextBox" runat="server" Width="32px"></asp:TextBox>
        :<asp:TextBox ID="MinTextBox" runat="server" Width="32px"></asp:TextBox>
        :<asp:TextBox ID="SecTextBox" runat="server" Width="32px"></asp:TextBox>
    (24 hour format)
        <p>&nbsp;</p>
    <h4>Meeting Time in UTC:<asp:TextBox ID="MeetingTimeTextBox" runat="server"></asp:TextBox>
    </h4>
    <asp:Button ID="TestButton" runat="server" Text="Test" />
    </li>
    </ol>
    <ol><li>
<<<<<<< HEAD
        <asp:DropDownList ID="DropDownList1" runat="server" DataSourceID="SqlDataSource1" DataTextField="Name" DataValueField="Name">
        </asp:DropDownList>
        <asp:DropDownList ID="DropDownList2" runat="server">
        </asp:DropDownList>
        <asp:Button ID="Button1" runat="server" Text="Update Meeting List" CausesValidation="false"/>
        <asp:SqlDataSource ID="SqlDataSource1" runat="server" ConnectionString="<%$ ConnectionStrings:UnoZoomCapstoneConnectionString %>" SelectCommand="SELECT [Name] FROM [Accounts]"></asp:SqlDataSource>
        </li></ol>
=======
        <asp:DropDownList ID="DropDownList1" runat="server"></asp:DropDownList>
        </li><li>
        <asp:Button ID="Button1" runat="server" Text="Button" CausesValidation="false"/></li></ol>
>>>>>>> origin/master
</asp:Content>
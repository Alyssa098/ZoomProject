﻿<%@ Page Title="Home Page" Language="VB" MasterPageFile="~/Site.Master" AutoEventWireup="true" CodeBehind="Default.aspx.vb" Inherits="BT_Zoom._Default" %>

<asp:Content runat="server" ID="FeaturedContent" ContentPlaceHolderID="FeaturedContent">
    <section class="featured">
        <div class="content-wrapper">
            <hgroup class="title">
                <h1><%: Title %>.</h1>
                <h2>Modify this template to jump-start your ASP.NET application.</h2>
            </hgroup>
            <p>
                To learn more about ASP.NET, visit <a href="http://asp.net" title="ASP.NET Website">http://asp.net</a>.
                The page features <mark>videos, tutorials, and samples</mark> to help you get the most from ASP.NET.
                If you have any questions about ASP.NET visit <a href="http://forums.asp.net/18.aspx" title="ASP.NET Forum">our forums</a>.
            </p>
        </div>
    </section>
</asp:Content>
<asp:Content runat="server" ID="BodyContent" ContentPlaceHolderID="MainContent">
    <h3>We suggest the following:</h3>
    <ol class="round">
        <li class="one">
            <h5>Getting Started</h5>
            ASP.NET Web Forms lets you build dynamic websites using a familiar drag-and-drop, event-driven model.
            A design surface and hundreds of controls and components let you rapidly build sophisticated, powerful UI-driven sites with data access.
            <a href="http://go.microsoft.com/fwlink/?LinkId=245146">Learn more…</a>
        </li>
        <li class="two">
            <h5>Add NuGet packages and jump-start your coding</h5>
            NuGet makes it easy to install and update free libraries and tools.
            <a href="http://go.microsoft.com/fwlink/?LinkId=245147">Learn more…</a>
        </li>
        <li class="three">
            <h5>Find Web Hosting</h5>
            You can easily find a web hosting company that offers the right mix of features and price for your applications.
            <a href="http://go.microsoft.com/fwlink/?LinkId=245143">Learn more…</a>
        </li>
    </ol>
    <ol>
    <li>
    <h4>Test Case from here</h4>
    <h4>Topic:<asp:TextBox ID="TopicTextBox" runat="server" Height="16px" Width="347px"></asp:TextBox>
        <asp:RequiredFieldValidator ID="RequiredFieldValidator2" runat="server" ControlToValidate="TopicTextBox" ErrorMessage="Topic is required" ForeColor="Red"></asp:RequiredFieldValidator>
    </h4>
    <h4>Meeing Date:<asp:Calendar ID="MeetingCalendar" runat="server"></asp:Calendar>
    
       <h4>Meeting Time</h4> 

    <asp:TextBox ID="HourTextBox" runat="server" Width="16px"></asp:TextBox>
        :<asp:TextBox ID="MinTextBox" runat="server" Width="16px"></asp:TextBox>
        :<asp:TextBox ID="SecTextBox" runat="server" Width="16px"></asp:TextBox>
    (24 hour format)
        <p>&nbsp;</p>
    <h4>Meeting Time in UTC:<asp:TextBox ID="MeetingTimeTextBox" runat="server"></asp:TextBox>
    </h4>
    <asp:Button ID="TestButton" runat="server" Text="Test" />
    </li>
    </ol>
    <ol><li>
        <asp:DropDownList ID="DropDownList1" runat="server"></asp:DropDownList>
        </li><li>
        <asp:Button ID="Button1" runat="server" Text="Button" CausesValidation="false"/></li></ol>
</asp:Content>
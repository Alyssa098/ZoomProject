<%@ Page Language="vb" AutoEventWireup="false"  MasterPageFile="~/Site.Master" CodeBehind="Update.aspx.vb" Inherits="BT_Zoom.Update" %>

<asp:Content runat="server" ID="FeaturedContent" ContentPlaceHolderID="FeaturedContent">
    <section class="featured">
        <div class="content-wrapper">
            <hgroup class="title">
                <h1><%: Title %></h1>
                <h2>Update Meeting</h2>
            </hgroup>
          
        </div>
    </section>
</asp:Content>
<asp:Content runat="server" ID="BodyContent" ContentPlaceHolderID="MainContent">
   
       
        
    
   
   

    <h4>
            &nbsp;</h4>
    <h4>Topic:<asp:TextBox ID="TopicTextBox" runat="server" Height="30px" Width="347px"></asp:TextBox>
        <asp:RequiredFieldValidator ID="RequiredFieldValidator2" runat="server" ControlToValidate="TopicTextBox" ErrorMessage="Topic is required" ForeColor="Red"></asp:RequiredFieldValidator>
    </h4>
    <h4>Meeing Date:<asp:Calendar ID="MeetingCalendar" runat="server"></asp:Calendar>
    
       <h4>Meeting Time</h4> 

    <asp:TextBox ID="HourTextBox" runat="server" Width="16px"></asp:TextBox>
        :<asp:TextBox ID="MinTextBox" runat="server" Width="16px"></asp:TextBox>
        :<asp:TextBox ID="SecTextBox" runat="server" Width="16px"></asp:TextBox>
    (24 hour format)
    
    </li>
    <p>
        <asp:Button ID="UpdateButton" runat="server" Text="Update" />
        <asp:Button ID="DeleteButton" runat="server" Text="Delete" CausesValidation="False" ValidateRequestMode="Disabled" />
        <asp:Button ID="Button1" runat="server" Text="Playback"  CausesValidation="False"/>
    </p>
    <p>
        &nbsp;</p>
</asp:Content>
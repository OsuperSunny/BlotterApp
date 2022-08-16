<%@ Page Title="Home Page" Language="C#" MasterPageFile="~/Site.Master" AutoEventWireup="true" CodeBehind="Default.aspx.cs" Inherits="TradeBlotter._Default" %>

<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
     <h1>Blotter</h1>
    <div class="jumbotron">
       
      
    <asp:Button ID="btnUpload" runat="server" OnClick="btnUpload_Click" Style="position: absolute; z-index:1;
top: 152px; left: 590px; height: 22px;right: 369px; width:65px" Text="Upload" TabIndex="11" Font-Bold="true"/>
        <asp:FileUpload ID="FileUpload2" runat="server" ToolTip="select file to upload" Style="z-index: 1; left: 288px; top: 151px;
        position: absolute; right: 446px;" TabIndex="10"/>
    </div>

  

</asp:Content>

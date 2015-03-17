<%@ Page Title="Distance" Language="VB" MasterPageFile="~/Site.master" AutoEventWireup="false" CodeFile="Distance.aspx.vb" Inherits="Default7" %>

<asp:Content ID="Content1" ContentPlaceHolderID="HeadContent" Runat="Server">
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="MainContent" Runat="Server">
    <p>
        This app allows you to calculate distance.</p>
    <p>
        <asp:ListBox ID="ListBox1" runat="server" DataSourceID="AccessDataSource1" DataTextField="city" DataValueField="ID"></asp:ListBox>
&nbsp;&nbsp;&nbsp;
        <asp:ListBox ID="ListBox2" runat="server" DataSourceID="AccessDataSource1" DataTextField="CITY" DataValueField="ID"></asp:ListBox>
        <asp:AccessDataSource ID="AccessDataSource1" runat="server" DataFile="~/App_Data/Distance.accdb" SelectCommand="SELECT * FROM [Table1]"></asp:AccessDataSource>
    </p>
    <p>
        <asp:Button ID="Button1" runat="server" Text="Calculate" />
&nbsp;&nbsp;&nbsp;
        <asp:Label ID="Label1" runat="server" Text="Label"></asp:Label>
    </p>
    <p>
        &nbsp;</p>
</asp:Content>


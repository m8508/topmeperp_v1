<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="ComparisionForm.aspx.cs" Inherits="topmeperp.Views.Inquiry.ComparisionForm" %>
    <form id="form1" runat="server">
        <h3>
        <asp:Label ID="labelMsg" runat="server" Text="labelMsg"></asp:Label></h3>
        <asp:GridView ID="grdRawData" runat="server" CssClass="table table-bordered">
            <RowStyle BackColor="#F7F7DE" />
            <FooterStyle BackColor="#CCCC99" />
            <PagerStyle BackColor="#F7F7DE" ForeColor="Black" HorizontalAlign="Right" />
            <SelectedRowStyle BackColor="#CE5D5A" Font-Bold="True" ForeColor="White" />
            <HeaderStyle BackColor="#6B696B" Font-Bold="True" ForeColor="White" />
            <AlternatingRowStyle BackColor="White" />
        </asp:GridView>
    </form>
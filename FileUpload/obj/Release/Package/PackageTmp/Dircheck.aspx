<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="Dircheck.aspx.cs" Inherits="FileUpload.Dircheck" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
    <asp:button ID="btnImport" runat="server" text="Import" OnClick="btnImport_Click" />
         <asp:button ID="btnExport" runat="server" text="Export" OnClick="btnExport_Click" />
    </div>
    </form>
</body>

</html>

<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="Default.aspx.cs" Inherits="FileUpload.Default" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
    <script  type="text/javascript">
        function GetFile() {
            var id = document.getElementById("FileUploadMS").value;
            alert(id);
        }

    </script>
</head>
<body>
    <form id="form1" runat="server">
    <div>
        <asp:FileUpload ID="FileUploadMS" runat="server" OnDataBinding="FileUploadMS_DataBinding" />
        <asp:Button ID="btnExport" runat="server" 
               Text="Export" 
               style="width:85px" OnClick="btnExport_Click" OnClientClick ="GetFile();" />
   <br /><br />
   <asp:Label ID="lblmessage" runat="server" />

    </div>
    </form>
</body>
</html>

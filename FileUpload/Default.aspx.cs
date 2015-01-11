using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
//using Microsoft.Office.Interop;
using Microsoft.Office.Interop.Word;
using Microsoft.Office.Interop.Excel;
using System.IO;
using System.Security.AccessControl;

namespace FileUpload
{
    public partial class Default : System.Web.UI.Page
    {
       


        protected void Page_Load(object sender, EventArgs e)
        {

            lblmessage.Text = "";
           
        }


        private void WriteToFile(string strPath, ref byte[] Buffer)
        {
            
            FileStream newFile = new FileStream(strPath, FileMode.Create, FileAccess.ReadWrite, FileShare.ReadWrite);
            newFile.Write(Buffer, 0, Buffer.Length);
            newFile.Close();
        }

        private void FileWordConvert(string uploadPath)
        {
            string dir = Server.MapPath("~") + @"\Export";
            Microsoft.Office.Interop.Word.Application wordApplication;
            wordApplication = new Microsoft.Office.Interop.Word.Application();
            string filepath = Path.GetFullPath(FileUploadMS.PostedFile.FileName);        
            Document wordDocument = null;
            object paramMissing = Type.Missing;

            try
            {              
                //object paramSourceDocPath = Path.GetFullPath(FileUploadMS.PostedFile.FileName);
                object paramSourceDocPath = uploadPath + @"\" +FileUploadMS.FileName;
                string filename = FileUploadMS.FileName;
                string [] arrFile = filename.Split('.');
                string paramExportFilePath = dir + @"\" + arrFile[0] + ".pdf";            
                WdExportFormat paramExportFormat = WdExportFormat.wdExportFormatPDF;

                wordDocument = wordApplication.Documents.Open(paramSourceDocPath);
                wordDocument.ExportAsFixedFormat(paramExportFilePath, paramExportFormat);               
                wordDocument.Close();
                wordApplication.Quit();
            }
            catch (Exception ex)
            {
                lblmessage.Text = ex.Message;
            }
            
        
        }

        private void FileExCelConvert(string uploadPath)
        {
            try
            {
                string dir = Server.MapPath("~")+ @"\Export";
           
                var excelApp = new Microsoft.Office.Interop.Excel.Application();
                XlFixedFormatType paramExportFormat =  Microsoft.Office.Interop.Excel.XlFixedFormatType.xlTypePDF;
                //string filepath = Path.GetFullPath(FileUploadMS.PostedFile.FileName);
                string filepath = uploadPath + @"\" + FileUploadMS.FileName;
                var excelDocument = excelApp.Workbooks.Open(filepath);

                string filename = FileUploadMS.FileName;
                string[] arrFile = filename.Split('.');

                string paramExportFilePath = dir +@"\" + arrFile[0] + ".pdf";
                excelDocument.ExportAsFixedFormat(paramExportFormat,paramExportFilePath);
                excelDocument.Close(false, "", false); 
                excelApp.Quit();

            }
            catch (Exception ex)
            {
                lblmessage.Text = ex.Message;
            }
        }

        protected void FileUploadMS_DataBinding(object sender, EventArgs e)
        {
            
        }

        protected void btnExport_Click(object sender, EventArgs e)
        {

           string dirimport = Server.MapPath("~")+ @"Import";
            if (!Directory.Exists(dirimport))
           {

               Directory.CreateDirectory(dirimport);
               //givepermission("Import");
             
           
           }
            FolderPermission("IIS_IUSRS", dirimport);
            HttpPostedFile myFile = FileUploadMS.PostedFile;
            int nFileLen = myFile.ContentLength;
            
            byte[] myData = new byte[nFileLen];    
            myFile.InputStream.Read(myData, 0, nFileLen);
            WriteToFile(dirimport + @"\" + FileUploadMS.FileName , ref myData);
          
           string dir = Server.MapPath("~")+ @"\Export";
          
           if (! Directory.Exists(dir))
           {
           
                Directory.CreateDirectory(dir);
           
           }

           FolderPermission("IIS_IUSRS", dir);
 
            if (FileUploadMS.HasFile)
            {

                string filename = FileUploadMS.FileName.ToLower();
                try
                {
                    if (filename.Contains(".xls") || filename.Contains(".csv"))
                    {
                        FileExCelConvert(dirimport);
                        lblmessage.Text = "File uploaded successfully !";
                    }

                    else if (filename.Contains(".doc"))
                    {
                        FileWordConvert(dirimport);
                        lblmessage.Text = "File uploaded successfully !";
                    }

                    else 

                    {
                        lblmessage.Text = "File Format is not correct !";
                    
                    }
                    
                }

                catch (Exception ex)
                {

                    lblmessage.Text = ex.Message;

                }
            }
        }

        public void FolderPermission(String accountName, String folderPath)
        {
            try
            {

                FileSystemRights Rights;

                //What rights are we setting? Here accountName is == "IIS_IUSRS"


                lblmessage.Text = "FolderPermission";
                Rights = FileSystemRights.FullControl;
                bool modified;
                var none = new InheritanceFlags();
                none = InheritanceFlags.None;

                //set on dir itself
                var accessRule = new FileSystemAccessRule(accountName, Rights, none, PropagationFlags.NoPropagateInherit, AccessControlType.Allow);
                var dInfo = new DirectoryInfo(folderPath);
                var dSecurity = dInfo.GetAccessControl();
                dSecurity.ModifyAccessRule(AccessControlModification.Set, accessRule, out modified);

                //Always allow objects to inherit on a directory 
                var iFlags = new InheritanceFlags();
                iFlags = InheritanceFlags.ContainerInherit | InheritanceFlags.ObjectInherit;

                //Add Access rule for the inheritance
                var accessRule2 = new FileSystemAccessRule(accountName, Rights, iFlags, PropagationFlags.InheritOnly, AccessControlType.Allow);
                dSecurity.ModifyAccessRule(AccessControlModification.Add, accessRule2, out modified);

                dInfo.SetAccessControl(dSecurity);
                lblmessage.Text = "FolderPermission";
            }
            catch (Exception ex)
            {
                //MessageBox.Show("Error");
                lblmessage.Text = ex.Message ;
            }
        }
        
        public static void AddDirectorySecurity(string FileName, string Account, FileSystemRights Rights, AccessControlType ControlType)
        {
            DirectoryInfo dInfo = new DirectoryInfo(FileName);
            DirectorySecurity dSecurity = dInfo.GetAccessControl();
            dSecurity.AddAccessRule(
                new System.Security.AccessControl.FileSystemAccessRule("IIS_IUSRS", Rights, ControlType));
            dInfo.SetAccessControl(dSecurity);
        }

        private void givepermission(string path)
        {
            DirectoryInfo a = new DirectoryInfo(Server.MapPath("~/" + path));
            AddDirectorySecurity(Server.MapPath("~/"), "IUSR", FileSystemRights.FullControl, AccessControlType.Allow);
        }

    }
}
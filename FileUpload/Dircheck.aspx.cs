using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace FileUpload
{
    public partial class Dircheck : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
           
        }

        protected void btnImport_Click(object sender, EventArgs e)
        {
            string dirimport = Server.MapPath("~") + @"Import";
            System.Security.AccessControl.FileSecurity sec = System.IO.File.GetAccessControl(dirimport);
            Response.Write("Owner : " + sec.GetOwner(typeof(System.Security.Principal.NTAccount)).Value);
            System.Security.AccessControl.AuthorizationRuleCollection auth = sec.GetAccessRules(true, true, typeof(System.Security.Principal.NTAccount));

            foreach (System.Security.AccessControl.FileSystemAccessRule objR in auth)
            {
                Response.Write(objR.IdentityReference);
            }
        }

        protected void btnExport_Click(object sender, EventArgs e)
        {
            string dirimport = Server.MapPath("~") + @"Export";
            System.Security.AccessControl.FileSecurity sec = System.IO.File.GetAccessControl(dirimport);
            Response.Write("Owner : " + sec.GetOwner(typeof(System.Security.Principal.NTAccount)).Value);
            System.Security.AccessControl.AuthorizationRuleCollection auth = sec.GetAccessRules(true, true, typeof(System.Security.Principal.NTAccount));

            foreach (System.Security.AccessControl.FileSystemAccessRule objR in auth)
            {
                Response.Write(objR.IdentityReference);
            }
        }
    }
}
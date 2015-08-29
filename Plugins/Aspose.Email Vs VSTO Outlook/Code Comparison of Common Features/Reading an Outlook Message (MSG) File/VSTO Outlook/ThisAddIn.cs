using Microsoft.Office.Interop.Outlook;
using System;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace VSTO_Outlook
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            // Create a new Application Class
            _Application outlook = new Outlook.Application();
            // Create a MailItem object
            MailItem item = (MailItem)outlook.CreateItemFromTemplate("test.msg", Type.Missing);
            // Access different fields of the message
            System.Console.WriteLine(string.Format("Subject:{0}", item.Subject));
            System.Console.WriteLine(string.Format("Sender Email Address:{0}", item.SenderEmailAddress));
            System.Console.WriteLine(string.Format("SenderName:{0}", item.SenderName));
            System.Console.WriteLine(string.Format("TO:{0}", item.To));
            System.Console.WriteLine(string.Format("CC:{0}", item.CC));
            System.Console.WriteLine(string.Format("BCC:{0}", item.BCC));
            System.Console.WriteLine(string.Format("Html Body:{0}", item.HTMLBody));
            System.Console.WriteLine(string.Format("Text Body:{0}", item.Body));

        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}

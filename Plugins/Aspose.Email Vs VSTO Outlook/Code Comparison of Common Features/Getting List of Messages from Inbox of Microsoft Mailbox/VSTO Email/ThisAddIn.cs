using System;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace VSTO_Email
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            // Create Application class and get namespace
            Outlook.Application outlook = new Outlook.Application();
            Outlook.NameSpace ns = outlook.GetNamespace("Mapi");

            object _missing = Type.Missing;
            ns.Logon(_missing, _missing, false, true);

            // Get Inbox information in objec of type MAPIFolder
            Outlook.MAPIFolder inbox = ns.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);

            // Unread emails
            int unread = inbox.UnReadItemCount;

            // Display the subject of emails in the Inbox folder
            foreach (Outlook.MailItem mail in inbox.Items)
            {
                Console.WriteLine(mail.Subject);
            }
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

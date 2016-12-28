using Microsoft.Office.Interop.Outlook;
using System;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace VSTO_Outlook
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            string pstFilePath = "sample.pst";
            Outlook.Application app = new Application();
            NameSpace outlookNs = app.GetNamespace("MAPI");
            // Add PST file (Outlook Data File) to Default Profile
            outlookNs.AddStore(pstFilePath);
            MAPIFolder rootFolder = outlookNs.Stores["sample"].GetRootFolder();
            // Traverse through all folders in the PST file
            // TODO: This is not recursive
            Folders subFolders = rootFolder.Folders;
            foreach (Folder folder in subFolders)
            {
                Items items = folder.Items;
                foreach (object item in items)
                {
                    if (item is MailItem)
                    {
                        // Retrieve the Object into MailItem
                        MailItem mailItem = item as MailItem;
                        Console.WriteLine("Saving message {0} ....", mailItem.Subject);
                        // Save the message to disk in MSG format
                        // TODO: File name may contain invalid characters [\ / : * ? " < > |]
                        mailItem.SaveAs(@"\extracted\" + mailItem.Subject + ".msg", OlSaveAsType.olMSG);
                    }
                }
            }
            // Remove PST file from Default Profile
            outlookNs.RemoveStore(rootFolder);
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

using Outlook = Microsoft.Office.Interop.Outlook;

namespace VSTO_Email
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            // Creates a new Outlook Application instance
            Outlook.Application objOutlook = new Outlook.Application();

            // Creating a new Outlook message from the Outlook Application instance
            Outlook.MailItem msgInterop = (Outlook.MailItem)(objOutlook.CreateItem(Outlook.OlItemType.olMailItem));

            // Set recipient information
            msgInterop.To = "to@domain.com";
            msgInterop.CC = "cc@domain.com";

            // Set the message subject
            msgInterop.Subject = "Subject";

            // Set some HTML text in the HTML body
            msgInterop.HTMLBody = "<h3>HTML Heading 3</h3> <u>This is underlined text</u>";

            // Save the MSG file in local disk
            string strMsg = "TestInterop.msg";
            msgInterop.SaveAs(strMsg, Outlook.OlSaveAsType.olMSG);
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

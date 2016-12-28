using Outlook = Microsoft.Office.Interop.Outlook;

namespace VSTO_Email
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            // Create an object of type Outlook.Application
            Outlook.Application objOutlook = new Outlook.Application();

            //Create an object of type olMailItem
            Outlook.MailItem oIMailItem = objOutlook.CreateItem(Outlook.OlItemType.olMailItem);

            //Set properties of the message file e.g. subject, body and to address
            //Set subject
            oIMailItem.Subject = "This MSG file is created using Office Automation.";
            //Set to (recipient) address
            oIMailItem.To = "to@domain.com";
            //Set body of the email message
            oIMailItem.HTMLBody = "<html><p>This MSG file is created using VBA code.</p>";

            //Add attachments to the message
            oIMailItem.Attachments.Add("image.bmp");
            oIMailItem.Attachments.Add("pic.jpeg");

            //Save as Outlook MSG file
            oIMailItem.SaveAs("testvba.msg");

            //Open the MSG file
            oIMailItem.Display();
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

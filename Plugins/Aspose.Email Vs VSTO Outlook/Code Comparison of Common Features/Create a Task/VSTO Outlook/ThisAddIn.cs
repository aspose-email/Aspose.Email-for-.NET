using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;

namespace VSTO_Outlook
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            CreateToDoItemExample();
        }
        private void CreateToDoItemExample()
        {
            // Date operations
            DateTime today = DateTime.Parse("10:00 AM");
            TimeSpan duration = TimeSpan.FromDays(1);
            DateTime tomorrow = today.Add(duration);
            Outlook.MailItem mail = Application.Session.
                GetDefaultFolder(
                Outlook.OlDefaultFolders.olFolderInbox).Items.Find(
                "[MessageClass]='IPM.Note'") as Outlook.MailItem;
            mail.MarkAsTask(Outlook.OlMarkInterval.olMarkTomorrow);
            mail.TaskStartDate = today;
            mail.ReminderSet = true;
            mail.ReminderTime = tomorrow;
            mail.Save();
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

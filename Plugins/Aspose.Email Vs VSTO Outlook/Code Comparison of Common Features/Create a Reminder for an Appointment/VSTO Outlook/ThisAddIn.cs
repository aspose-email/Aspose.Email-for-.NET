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
            ReminderExample();
        }
        private void ReminderExample()
        {
            Outlook.AppointmentItem appt = Application.CreateItem(
                Outlook.OlItemType.olAppointmentItem)
                as Outlook.AppointmentItem;
            appt.Subject = "Wine Tasting";
            appt.Location = "Napa CA";
            appt.Sensitivity = Outlook.OlSensitivity.olPrivate;
            appt.Start = DateTime.Parse("10/21/2006 10:00 AM");
            appt.End = DateTime.Parse("10/21/2006 3:00 PM");
            appt.ReminderSet = true;
            appt.ReminderMinutesBeforeStart = 120;
            appt.Save();
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

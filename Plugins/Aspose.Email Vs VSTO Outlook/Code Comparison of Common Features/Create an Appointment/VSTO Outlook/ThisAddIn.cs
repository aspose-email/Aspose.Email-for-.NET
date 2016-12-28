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
            AllDayEventExample();
        }
        private void AllDayEventExample()
        {
            Outlook.AppointmentItem appt = Application.CreateItem(
                Outlook.OlItemType.olAppointmentItem)
                as Outlook.AppointmentItem;
            appt.Subject = "Developer's Conference";
            appt.AllDayEvent = true;
            appt.Start = DateTime.Parse("6/11/2007 12:00 AM");
            appt.End = DateTime.Parse("6/16/2007 12:00 AM");
            appt.Display(false);
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

using System;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace VSTO_Outlook
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            // Create an instance of Outlook Application class
            Outlook.Application outlookApp = new Outlook.Application();

            // Create an instance of AppointmentItem object and set the properties:
            Outlook.AppointmentItem oAppointment = (Outlook.AppointmentItem)outlookApp.CreateItem(Outlook.OlItemType.olAppointmentItem);

            oAppointment.Subject = "subject of appointment";
            oAppointment.Body = "body text of appointment";
            oAppointment.Location = "Appointment location";

            // Set the start date and end dates
            oAppointment.Start = Convert.ToDateTime("01/22/2010 10:00:00 AM");
            oAppointment.End = Convert.ToDateTime("01/22/2010 2:00:00 PM");

            // Save the appointment
            oAppointment.Save();

            // Send the appointment
            Outlook.MailItem mailItem = oAppointment.ForwardAsVcal();
            mailItem.To = "recipient@domain.com";
            mailItem.Send();
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

using Aspose.Email.Calendar;
using Aspose.Email.Clients.Exchange.WebService;
using Aspose.Email.Mime;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Aspose.Email.Examples.EWS
{
    class CreatingUpdatingAndDeletingCalendarItemsUsingEWS
    {
        public static void Run()
        {
            try
            {
                IEWSClient client = EWSClient.GetEWSClient("https://outlook.office365.com/ews/exchange.asmx", "your.username", "your.Password");
                DateTime date = DateTime.Now;
                DateTime startTime = new DateTime(date.Year, date.Month, date.Day, date.Hour, 0, 0);
                DateTime endTime = startTime.AddHours(1);
                string timeZone = "America/New_York";

                Appointment app = new Appointment("Room 112", startTime, endTime, "organizeraspose-email.test3@domain.com","attendee@gmail.com");
                app.SetTimeZone(timeZone);
                app.Summary = "NETWORKNET-34136" + Guid.NewGuid().ToString();
                app.Description = "NETWORKNET-34136 Exchange 2007/EWS: Provide support for Add/Update/Delete calendar items";

                string uid = client.CreateAppointment(app);
                Appointment fetchedAppointment1 = client.FetchAppointment(uid);
                app.Location = "Room 115";
                app.Summary = "New summary for " + app.Summary;
                app.Description = "New Description";
                client.UpdateAppointment(app);

                Appointment[] appointments1 = client.ListAppointments();
                Console.WriteLine("Total Appointments: "  + appointments1.Length);
                Appointment fetchedAppointment2 = client.FetchAppointment(uid);
                Console.WriteLine("Summary: " + fetchedAppointment2.Summary);
                Console.WriteLine("Location: " + fetchedAppointment2.Location);
                Console.WriteLine("Description: " + fetchedAppointment2.Description);
                client.CancelAppointment(app);
                Appointment[] appointments2 = client.ListAppointments();
                Console.WriteLine("Total Appointments: " + appointments2.Length);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
    }
}

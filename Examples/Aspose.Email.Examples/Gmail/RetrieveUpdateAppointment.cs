using Aspose.Email.Calendar;
using Aspose.Email.Clients.Google;
using Aspose.Email.Mime;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Aspose.Email.Examples.Gmail
{
    class RetrieveUpdateAppointment
    {
        public static void Run()
        {
            try
            {
                GoogleTestUser User2 = new GoogleTestUser("user", "email address", "password", "clientId", "client secret");
                string accessToken;
                string refreshToken;
                GoogleOAuthHelper.GetAccessToken(User2, out accessToken, out refreshToken);
                             
                // Get IGmailclient
                using (IGmailClient client = GmailClient.GetInstance(accessToken, User2.EMail))
                {
                    string calendarId = client.ListCalendars()[0].Id;
                    string AppointmentUniqueId = client.ListAppointments(calendarId)[0].UniqueId;

                    // Retrieve Appointment
                    Appointment app3 = client.FetchAppointment(calendarId, AppointmentUniqueId);
                    // Change the appointment information
                    app3.Summary = "New Summary - " + Guid.NewGuid().ToString();
                    app3.Description = "New Description - " + Guid.NewGuid().ToString();
                    app3.Location = "New Location - " + Guid.NewGuid().ToString();
                    app3.Flags = AppointmentFlags.AllDayEvent;
                    app3.StartDate = DateTime.Now.AddHours(2);
                    app3.EndDate = app3.StartDate.AddHours(1);
                    app3.StartTimeZone = "Europe/Kiev";
                    app3.EndTimeZone = "Europe/Kiev";
                    // Update the appointment and get back updated appointment
                    Appointment app4 = client.UpdateAppointment(calendarId, app3);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
    }
}

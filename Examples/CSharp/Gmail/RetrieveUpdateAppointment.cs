using Aspose.Email.Calendar;
using Aspose.Email.Clients.Google;
using Aspose.Email.Mime;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/

namespace Aspose.Email.Examples.CSharp.Email.Gmail
{
    class RetrieveUpdateAppointment
    {
        public static void Run()
        {
            try
            {
                // ExStart:RetrieveUpdateAppointment
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
                // ExEnd:RetrieveUpdateAppointment
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
    }
}

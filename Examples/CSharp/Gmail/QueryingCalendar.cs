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
    class QueryingCalendar
    {
        public static void Run()
        {
            try
            {
                // ExStart:QueryingCalendar
                // Get access token
                GoogleTestUser User2 = new GoogleTestUser("user", "email address", "password", "clientId", "client secret");
                string accessToken;
                string refreshToken;
                GoogleOAuthHelper.GetAccessToken(User2, out accessToken, out refreshToken);

                // Get IGmailClient
                using (IGmailClient client = GmailClient.GetInstance(accessToken, User2.EMail))
                {
                    // Initialize calendar item
                    Aspose.Email.Clients.Google.Calendar calendar1 = new Aspose.Email.Clients.Google.Calendar("summary - " + Guid.NewGuid().ToString(), null, null, "Europe/Kiev");

                    // Insert calendar and get back id of newly inserted calendar and Fetch the same calendar using calendar id
                    string id = client.CreateCalendar(calendar1);
                    Aspose.Email.Clients.Google.Calendar cal1 = client.FetchCalendar(id);
                    string calendarId1 = cal1.Id;
                    try
                    {
                        // Get list of appointments in newly inserted calendar. It should be zero
                        Appointment[] appointments = client.ListAppointments(calendarId1);
                        if (appointments.Length != 0)
                        {
                            Console.WriteLine("Wrong number of appointments");
                            return;
                        }

                        // Create a new appointment and Calculate appointment start and finish time
                        DateTime startDate = DateTime.Now;
                        DateTime endDate = startDate.AddHours(1);

                        // Create attendees list for appointment
                        MailAddressCollection attendees = new MailAddressCollection();
                        attendees.Add("user1@domain.com");
                        attendees.Add("user2@domain.com");

                        // Create appointment
                        Appointment app1 = new Appointment("Location - " + Guid.NewGuid().ToString(), startDate, endDate, "user2@domain.com", attendees);
                        app1.Summary = "Summary - " + Guid.NewGuid().ToString();
                        app1.Description = "Description - " + Guid.NewGuid().ToString();
                        app1.StartTimeZone = "Europe/Kiev";
                        app1.EndTimeZone = "Europe/Kiev";

                        // Insert the newly created appointment and get back the same in case of successful insertion
                        Appointment app2 = client.CreateAppointment(calendarId1, app1);

                        // Create Freebusy query by setting min/max timeand time zone
                        FreebusyQuery query = new FreebusyQuery();
                        query.TimeMin = DateTime.Now.AddDays(-1);
                        query.TimeMax = DateTime.Now.AddDays(1);
                        query.TimeZone = "Europe/Kiev";

                        // Set calendar item to search and Get the reponse of query containing 
                        query.Items.Add(cal1.Id);
                        FreebusyResponse resp = client.GetFreebusyInfo(query);
                        // Delete the appointment
                        client.DeleteAppointment(calendarId1, app2.UniqueId);
                    }
                    finally
                    {
                        // Delete the calendar
                        client.DeleteCalendar(cal1.Id);
                    }
                }
                // ExEnd:QueryingCalendar
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
                    
        }
    }
}

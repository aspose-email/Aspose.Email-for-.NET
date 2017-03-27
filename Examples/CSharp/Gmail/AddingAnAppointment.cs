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
    class AddingAnAppointment
    {
        public static void Run()
        {
            try
            {
                // ExStart:AddingAnAppointment

                GoogleTestUser User2 = new GoogleTestUser("user", "email address", "password", "clientId", "client secret");
                string accessToken;
                string refreshToken;
                GoogleOAuthHelper.GetAccessToken(User2, out accessToken, out refreshToken);

                // Get IGmailclient
                using (IGmailClient client = GmailClient.GetInstance(accessToken, User2.EMail))
                {
                    // Create local calendar	
                    Aspose.Email.Clients.Google.Calendar calendar1 = new Aspose.Email.Clients.Google.Calendar("summary - " + Guid.NewGuid().ToString(), null, null, "Europe/Kiev");

                    // Insert calendar and get id of inserted calendar and Get back calendar using an id
                    string id = client.CreateCalendar(calendar1);
                    Aspose.Email.Clients.Google.Calendar cal1 = client.FetchCalendar(id);
                    string calendarId1 = cal1.Id;

                    try
                    {
                        // Retrieve list of appointments from the first calendar
                        Appointment[] appointments = client.ListAppointments(calendarId1);
                        if (appointments.Length > 0)
                        {
                            Console.WriteLine("Wrong number of appointments");
                            return;
                        }

                        // Get current time and Calculate time after an hour from now
                        DateTime startDate = DateTime.Now;
                        DateTime endDate = startDate.AddHours(1);

                        // Initialize a mail address collection and set attendees mail address
                        MailAddressCollection attendees = new MailAddressCollection();
                        attendees.Add("User1.EMail@domain.com");
                        attendees.Add("User3.EMail@domain.com");

                        // Create an appointment with above attendees
                        Appointment app1 = new Appointment("Location - " + Guid.NewGuid().ToString(), startDate, endDate, User2.EMail, attendees);

                        // Set appointment summary, description, start/end time zone
                        app1.Summary = "Summary - " + Guid.NewGuid().ToString();
                        app1.Description = "Description - " + Guid.NewGuid().ToString();
                        app1.StartTimeZone = "Europe/Kiev";
                        app1.EndTimeZone = "Europe/Kiev";

                        // Insert appointment in the first calendar inserted above and get back inserted appointment
                        Appointment app2 = client.CreateAppointment(calendarId1, app1);

                        // Retrieve appointment using unique id
                        Appointment app3 = client.FetchAppointment(calendarId1, app2.UniqueId);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }
                }
                // ExEnd:AddingAnAppointment
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
    }
}

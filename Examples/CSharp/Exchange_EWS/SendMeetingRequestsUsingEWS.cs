using System;
using System.Net;
using Aspose.Email.Clients.Exchange.WebService;
using Aspose.Email.Mime;
using Aspose.Email.Calendar;
using Aspose.Email.Calendar.Recurrences;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/

namespace Aspose.Email.Examples.CSharp.Email.Exchange_EWS
{
    class SendMeetingRequestsUsingEWS
    {
        public static void Run()
        {
            // ExStart:SendMeetingRequestsUsingEWS
            try
            {
                // Create instance of IEWSClient class by giving credentials
                IEWSClient client = EWSClient.GetEWSClient("https://outlook.office365.com/ews/exchange.asmx", "testUser", "pwd", "domain");

                // Create the meeting request
                Appointment app = new Appointment("meeting request", DateTime.Now.AddHours(1), DateTime.Now.AddHours(1.5), "administrator@test.com", "bob@test.com");
                app.Summary = "meeting request summary";
                app.Description = "description";


                RecurrencePattern pattern = new Aspose.Email.Calendar.Recurrences.DailyRecurrencePattern(DateTime.Now.AddDays(5));
                app.Recurrence = pattern;

                // Create the message and set the meeting request
                MailMessage msg = new MailMessage();
                msg.From = "administrator@test.com";
                msg.To = "bob@test.com";
                msg.IsBodyHtml = true;
                msg.HtmlBody = "<h3>HTML Heading</h3><p>Email Message detail</p>";
                msg.Subject = "meeting request";
                msg.AddAlternateView(app.RequestApointment(0));

                // send the appointment
                client.Send(msg);
                Console.WriteLine("Appointment request sent");
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            // ExEnd:SendMeetingRequestsUsingEWS
        }        
    }
}
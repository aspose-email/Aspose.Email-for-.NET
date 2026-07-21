using System;
using System.Net;
using Aspose.Email.Clients.Exchange.WebService;
using Aspose.Email.Mime;
using Aspose.Email.Calendar;
using Aspose.Email.Calendar.Recurrences;

namespace Aspose.Email.Examples.EWS
{
    class SendMeetingRequestsUsingEWS
    {
        public static void Run()
        {
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
        }        
    }
}
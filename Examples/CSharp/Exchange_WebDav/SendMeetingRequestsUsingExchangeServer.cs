using System;
using System.Net;
using Aspose.Email.Clients.Exchange.Dav;
using Aspose.Email.Mime;
using Aspose.Email.Calendar;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/

namespace Aspose.Email.Examples.CSharp.Email.Exchange_WebDav
{
    class SendMeetingRequestsUsingExchangeServer
    {
        public static void Run()
        {
            // ExStart:SendMeetingRequestsUsingExchangeServer
            try
            {
                string mailboxUri = @"https://ex07sp1/exchange/administrator"; // WebDAV
                string domain = @"litwareinc.com";
                string username = @"administrator";
                string password = @"Evaluation1";

                NetworkCredential credential = new NetworkCredential(username, password, domain);
                Console.WriteLine("Connecting to Exchange server.....");
                // Connect to Exchange Server
                ExchangeClient client = new ExchangeClient(mailboxUri, credential); // WebDAV

                // Create the meeting request
                Appointment app = new Appointment("meeting request", DateTime.Now.AddHours(1), DateTime.Now.AddHours(1.5), "administrator@" + domain, "bob@" + domain);
                app.Summary = "meeting request summary";
                app.Description = "description";

                // Create the message and set the meeting request
                MailMessage msg = new MailMessage();
                msg.From = "administrator@" + domain;
                msg.To = "bob@" + domain;
                msg.IsBodyHtml = true;
                msg.HtmlBody = "<h3>HTML Heading</h3><p>Email Message detail</p>";
                msg.Subject = "meeting request";
                msg.AddAlternateView(app.RequestApointment(0));

                // Send the appointment
                client.Send(msg);
                Console.WriteLine("Appointment request sent");
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            // ExEnd:SendMeetingRequestsUsingExchangeServer
        }        
    }
}
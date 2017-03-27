using System;
using System.Diagnostics;
using Aspose.Email.Mime;
using Aspose.Email.Clients.Smtp;
using Aspose.Email.Clients;
using Aspose.Email.Calendar;

namespace Aspose.Email.Examples.CSharp.Email.IMAP
{
    class MeetingRequests
    {
        public static void Run()
        {
            // The path to the File directory.
            string dataDir = RunExamples.GetDataDir_SMTP();
            string dstEmail = dataDir + "outputAttachments.msg";

            // ExStart:SendingMeetingRequestsViaEmail
            // Create an instance of the MailMessage class
            MailMessage msg = new MailMessage();

            // Set the sender, recipient, who will receive the meeting request. Basically, the recipient is the same as the meeting attendees
            msg.From = "newcustomeronnet@gmail.com";
            msg.To = "person1@domain.com, person2@domain.com, person3@domain.com, asposetest123@gmail.com";

            // Create Appointment instance
            Appointment app = new Appointment("Room 112", new DateTime(2015, 7, 17, 13, 0, 0), new DateTime(2015, 7, 17, 14, 0, 0), msg.From, msg.To);
            app.Summary = "Release Meetting";
            app.Description = "Discuss for the next release";

            // Add appointment to the message and Create an instance of SmtpClient class
            msg.AddAlternateView(app.RequestApointment());
            SmtpClient client = GetSmtpClient();

            try
            {
                // Client.Send will send this message
                client.Send(msg);
                Console.WriteLine("Message sent");
            }
            catch (Exception ex)
            {
                Trace.WriteLine(ex.ToString());
            }
            // ExEnd:SendingMeetingRequestsViaEmail

            Console.WriteLine(Environment.NewLine + "Meeting request send successfully.");
        }

        private static SmtpClient GetSmtpClient()
        {
            SmtpClient client = new SmtpClient("smtp.gmail.com", 587, "your.email@gmail.com", "your.password");
            client.SecurityOptions = SecurityOptions.Auto;
            return client;
        }
    }
}
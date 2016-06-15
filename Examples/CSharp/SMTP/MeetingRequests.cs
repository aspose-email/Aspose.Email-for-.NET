using System.IO;
using System;
using Aspose.Email.Mail;
using Aspose.Email.Outlook;
using Aspose.Email.Pop3;
using Aspose.Email;
using Aspose.Email.Mime;
using Aspose.Email.Imap;
using System.Configuration;
using System.Data;

namespace Aspose.Email.Examples.CSharp.SMTP
{
    class MeetingRequests
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_SMTP();
            string dstEmail = dataDir + "outputAttachments.msg";

            //Create an instance of the MailMessage class
            MailMessage msg = new MailMessage();

            //set the sender
            msg.From = "newcustomeronnet@gmail.com";

            //Set the recipient, who will receive the meeting request.
            //Basically, the recipient is the same as the meeting attendees.
            msg.To = "person1@domain.com, person2@domain.com, person3@domain.com, asposetest123@gmail.com";

            //create Appointment instance
            Appointment app = new Appointment(
            "Room 112", //location
            new DateTime(2015, 7, 17, 13, 0, 0), //start time
            new DateTime(2015, 7, 17, 14, 0, 0), //end time
            msg.From,//organizer
            msg.To //attendee
            );

            app.Summary = "Release Meetting";
            app.Description = "Discuss for the next release";

            //add appointment to the message
            msg.AddAlternateView(app.RequestApointment());

            //Create an instance of SmtpClient class
            SmtpClient client = GetSmtpClient();

            try
            {
                //Client.Send will send this message
                client.Send(msg);
                //Show Message Sent�E only if message sent successfully
                Console.WriteLine("Message sent");
            }

            catch (Exception ex)
            {
                System.Diagnostics.Trace.WriteLine(ex.ToString());
            }

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

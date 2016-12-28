using Aspose.Email.Mail;
using System;

namespace Aspose_Email
{
    class Program
    {
        static void Main(string[] args)
        {
            // Create attendees of the meeting
            MailAddressCollection attendees = new MailAddressCollection();
            attendees.Add("recipient1@domain.com");
            attendees.Add("recipient2@domain.com");

            // Set up appointment
            Appointment app = new Appointment(
                "Location", // location of meeting
                DateTime.Now, // start date
                DateTime.Now.AddHours(1), // end date
                new MailAddress("organizer@domain.com"), // organizer
                attendees); // attendees

            // Set up message that needs to be sent
            MailMessage msg = new MailMessage();
            msg.From = "from@domain.com";
            msg.To = "to@domain.com";
            msg.Subject = "appointment request";
            msg.Body = "you are invited";

            // Add meeting request to the message
            msg.AddAlternateView(app.RequestApointment());

            // Set up the SMTP client to send email with meeting request
            SmtpClient client = new SmtpClient("host", 25, "user", "password");
            client.Send(msg);
        }
    }
}

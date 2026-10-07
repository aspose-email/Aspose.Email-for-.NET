// Demonstrates how to send a meeting request.
//
// RequestApointment turns an Appointment into an iCalendar REQUEST part; added to a
// message as an alternate view, it makes Outlook, Gmail and other clients show the
// message as an invitation with Accept / Decline buttons. The attendees of the
// appointment are normally also the recipients of the message.
//
// The request is saved to Out. When SMTP is configured it is also sent, with your own
// address as the only attendee.

using System;
using Aspose.Email.Calendar;

namespace Aspose.Email.Examples.SMTP
{
    internal static class MeetingRequests
    {
        public static void Run()
        {
            var organizer = "organizer@example.com";
            var attendees = "attendee1@example.com, attendee2@example.com";

            using (var client = ClientBuilder.IsSmtpConfigured ? ClientBuilder.Smtp(AuthType.Basic) : null)
            {
                if (client != null)
                    organizer = attendees = client.Username;

                var start = DateTime.Today.AddDays(7).AddHours(13);
                var appointment = new Appointment("Room 112", start, start.AddHours(1), organizer, attendees)
                {
                    Summary = "Release meeting",
                    Description = "Let's discuss the next release."
                };

                var message = new MailMessage { From = organizer, To = attendees, Subject = appointment.Summary };
                message.AddAlternateView(appointment.RequestApointment());

                var outputPath = Data.Out/"MeetingRequests_out.eml";
                message.Save(outputPath, SaveOptions.DefaultEml);
                Console.WriteLine($"Meeting request for {start:g} saved to {outputPath}");

                if (client == null)
                {
                    Console.WriteLine("Set Smtp.HostName in clientsettings.json to also send it.");
                    return;
                }

                client.Send(message);
                Console.WriteLine($"Sent to {attendees}.");
            }
        }
    }
}

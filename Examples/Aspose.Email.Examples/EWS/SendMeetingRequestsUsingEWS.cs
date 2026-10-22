// Demonstrates sending a recurring meeting request.
//
// RequestApointment turns an Appointment into an iCalendar REQUEST part; added to a
// message as an alternate view, it makes Outlook show the message as an invitation with
// Accept / Decline buttons. A Recurrence pattern makes it a series - here daily for five
// days. The request goes to the configured mailbox itself.

using System;
using Aspose.Email.Calendar;
using Aspose.Email.Calendar.Recurrences;

namespace Aspose.Email.Examples.EWS
{
    internal static class SendMeetingRequestsUsingEWS
    {
        public static void Run()
        {
            var self = ClientBuilder.EwsUserName;
            var start = DateTime.Today.AddDays(1).AddHours(9);

            using (var client = ClientBuilder.Ews(AuthType.ModernWithAppPermission))
            {
                var appointment = new Appointment("Room 4", start, start.AddMinutes(15), self, self)
                {
                    Summary = "Daily stand-up",
                    Description = "Fifteen minutes, every morning this week.",
                    Recurrence = new DailyRecurrencePattern(start.AddDays(4))
                };

                var message = new MailMessage(self, self)
                {
                    Subject = appointment.Summary,
                    HtmlBody = "<p>Please join the daily stand-up.</p>"
                };
                message.AddAlternateView(appointment.RequestApointment());

                client.Send(message);
                Console.WriteLine($"Sent the meeting request '{appointment.Summary}' to {self}.");
            }
        }
    }
}

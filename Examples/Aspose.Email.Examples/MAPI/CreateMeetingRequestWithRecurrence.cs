// Demonstrates how to create a meeting request with a weekly recurrence pattern
// and send it via SMTP.

using System;
using Aspose.Email.Calendar;
using Aspose.Email.Calendar.Recurrences;

namespace Aspose.Email.Examples.MAPI
{
    internal static class CreateMeetingRequestWithRecurrence
    {
        public static void Run()
        {
            var msg = new MailMessage { From = new MailAddress("from@gmail.com") };
            msg.To.Add("to@domain.com");

            var startDate = new DateTime(2013, 12, 1, 17, 0, 0);
            var endDate = new DateTime(2013, 12, 31, 17, 30, 0);

            var appointment = new Appointment("same place", startDate, endDate, msg.From, msg.To)
            {
                UniqueId = Guid.NewGuid().ToString(),
                Description = "Recurring team meeting."
            };

            // Repeat every week on Monday, Tuesday and Thursday, for 14 occurrences.
            appointment.Recurrence = new WeeklyRecurrencePattern(14)
            {
                StartDays = new[] { CalendarDay.Monday, CalendarDay.Tuesday, CalendarDay.Thursday },
                Interval = 1
            };

            // RequestApointment() turns the appointment into the calendar part that mail
            // clients render as a meeting invitation.
            msg.AlternateViews.Add(appointment.RequestApointment());

            var outputPath = Data.Out/"CreateMeetingRequestWithRecurrence_out.eml";
            msg.Save(outputPath, SaveOptions.DefaultEml);

            Console.WriteLine($"Meeting request: {appointment.Location}, {startDate:g} - {endDate:g}");
            Console.WriteLine($"Recurrence:      {appointment.Recurrence}");
            Console.WriteLine($"Saved to:        {outputPath}");

            // Sending needs a real server, so it only runs once one is configured in
            // clientsettings.json. The meeting request above is complete either way.
            if (!ClientBuilder.IsSmtpConfigured)
            {
                Console.WriteLine("\nSet Smtp.HostName in clientsettings.json to also send this request.");
                return;
            }

            using (var client = ClientBuilder.Smtp(AuthType.Basic))
            {
                client.Send(msg);
                Console.WriteLine("\nMeeting request sent.");
            }
        }
    }
}

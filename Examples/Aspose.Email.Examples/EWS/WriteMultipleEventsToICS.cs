// Demonstrates writing several events into one iCalendar (.ics) file.
//
// CalendarWriter appends appointments to a single .ics file one by one, so a long
// list - for example everything read from an Exchange calendar - can be exported
// without holding it all in memory. The file imports into Outlook, Google Calendar and
// other calendar programs. This example runs offline.

using System;
using Aspose.Email.Calendar;

namespace Aspose.Email.Examples.EWS
{
    internal static class WriteMultipleEventsToICS
    {
        public static void Run()
        {
            var outputPath = Data.Out/"WriteMultipleEventsToICS_out.ics";
            var saveOptions = new AppointmentIcsSaveOptions { Action = AppointmentAction.Create };
            var start = DateTime.Today.AddDays(1).AddHours(9);

            using (var writer = new CalendarWriter(outputPath, saveOptions))
            {
                for (var i = 0; i < 5; i++)
                {
                    var appointment = new Appointment("Room 101", start.AddDays(i), start.AddDays(i).AddMinutes(30),
                        "organizer@example.com", "attendee@example.com")
                    {
                        Summary = $"Daily stand-up {i + 1}",
                        Description = "Fifteen minutes, standing."
                    };

                    writer.Write(appointment);
                }
            }

            Console.WriteLine($"Wrote 5 events to {outputPath}");
        }
    }
}

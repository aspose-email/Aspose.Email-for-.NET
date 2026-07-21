// Demonstrates how to create a MapiCalendar appointment and save it as ICS.

using System;
using Aspose.Email.Mapi;
using Aspose.Email.Calendar;

namespace Aspose.Email.Examples.MAPI
{
    internal static class CreateAndSaveCalendarItems
    {
        public static void Run()
        {
            var calendar = new MapiCalendar(
                "LAKE ARGYLE WA 6743",
                "Appointment",
                "This is a very important meeting :)",
                new DateTime(2012, 10, 2, 13, 0, 0),
                new DateTime(2012, 10, 2, 14, 0, 0));

            var outputPath = Data.Out/"CalendarItem_out.ics";
            calendar.Save(outputPath, AppointmentSaveFormat.Ics);

            Console.WriteLine($"{calendar.Subject} at {calendar.Location}");
            Console.WriteLine($"{calendar.StartDate:g} - {calendar.EndDate:t}");
            Console.WriteLine($"Saved to {outputPath}");
        }
    }
}

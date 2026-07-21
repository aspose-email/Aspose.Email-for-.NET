// Demonstrates how to create an Outlook appointment and save it as an MSG file.

using System;
using Aspose.Email.Calendar;
using Aspose.Email.Mapi;

namespace Aspose.Email.Examples.MAPI
{
    internal static class SavingTheCalendarItemAsMSG
    {
        public static void Run()
        {
            var startDate = new DateTime(2012, 10, 2, 13, 0, 0);
            var endDate = new DateTime(2012, 10, 2, 14, 0, 0);

            var calendar = new MapiCalendar(
                "LAKE ARGYLE WA 6743",
                "Appointment",
                "This is a very important meeting :)",
                startDate,
                endDate);

            var outputPath = Data.Out/"CalendarItemAsMSG_out.Msg";
            calendar.Save(outputPath, AppointmentSaveFormat.Msg);

            Console.WriteLine($"{calendar.Subject} at {calendar.Location}");
            Console.WriteLine($"{startDate:g} - {endDate:t}");
            Console.WriteLine($"Saved to {outputPath}");
        }
    }
}

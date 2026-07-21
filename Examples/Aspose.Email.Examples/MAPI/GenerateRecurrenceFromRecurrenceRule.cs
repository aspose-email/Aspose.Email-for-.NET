// Demonstrates how to create a MapiCalendar and assign a recurrence pattern from an RFC 5545 rule string.

using System;
using Aspose.Email.Mapi;
using Aspose.Email.Calendar;

namespace Aspose.Email.Examples.MAPI
{
    internal static class GenerateRecurrenceFromRecurrenceRule
    {
        public static void Run()
        {
            var startDate = new DateTime(2015, 7, 16);
            var endDate = new DateTime(2015, 8, 1);
            var app = new MapiCalendar("test location", "test summary", "test description", startDate, endDate);

            const string rule =
                "DTSTART;TZID=Europe/London:20150831T080000\r\n" +
                "DTEND;TZID=Europe/London:20150831T083000\r\n" +
                "RRULE:FREQ=DAILY;INTERVAL=1;COUNT=7\r\n" +
                "EXDATE:20150831T070000Z,20150904T070000Z";

            // The factory parses the rule text, so the pattern does not have to be built
            // property by property. EXDATE lists the occurrences to skip.
            app.Recurrence.RecurrencePattern = MapiCalendarRecurrencePatternFactory.FromString(rule);

            var outputPath = Data.Out/"GenerateRecurrenceFromRule_out.ics";
            app.Save(outputPath, AppointmentSaveFormat.Ics);

            var pattern = app.Recurrence.RecurrencePattern;
            Console.WriteLine($"Pattern:     {pattern.PatternType}, every {pattern.Period}");
            Console.WriteLine($"Occurrences: {pattern.OccurrenceCount}");
            Console.WriteLine($"Skipped:     {pattern.DeletedInstanceDates.Count} date(s)");
            Console.WriteLine($"Saved to {outputPath}");
        }
    }
}

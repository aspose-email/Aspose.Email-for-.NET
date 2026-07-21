// Demonstrates how to give an Outlook task a yearly recurrence that stops on a
// given date.

using System;
using Aspose.Email.Calendar.Recurrences;
using Aspose.Email.Mapi;

namespace Aspose.Email.Examples.MAPI
{
    internal static class YearlyEndAfterDate
    {
        public static void Run()
        {
            // Outlook stores task dates in UTC, so shift the local dates by the local
            // offset to keep the task on the intended day.
            var offset = TimeZoneInfo.Local.GetUtcOffset(DateTime.Now);

            var startDate = new DateTime(2015, 7, 1).Add(offset);
            var dueDate = new DateTime(2015, 7, 1).Add(offset);
            var endByDate = new DateTime(2020, 12, 31).Add(offset);

            var task = new MapiTask("This is test task", "Sample Body", startDate, dueDate)
            {
                State = MapiTaskState.NotAssigned
            };

            // A yearly pattern is a monthly one with a period of 12 months.
            var recurrence = new MapiCalendarMonthlyRecurrencePattern
            {
                Day = 15,
                Period = 12,
                PatternType = MapiCalendarRecurrencePatternType.Month,
                EndType = MapiCalendarRecurrenceEndType.EndAfterDate,
                EndDate = endByDate,
                OccurrenceCount = GetOccurrenceCount(startDate, endByDate, "FREQ=YEARLY;BYMONTHDAY=15;BYMONTH=7;INTERVAL=1")
            };

            // Outlook rejects a pattern with no occurrences at all.
            if (recurrence.OccurrenceCount == 0)
                recurrence.OccurrenceCount = 1;

            task.Recurrence = recurrence;

            var outputPath = Data.Out/"YearlyEndAfterDate_out.msg";
            task.Save(outputPath, TaskSaveFormat.Msg);

            Console.WriteLine($"Task recurs yearly until {endByDate:d}, {recurrence.OccurrenceCount} occurrence(s).");
            Console.WriteLine($"Saved to {outputPath}");
        }

        // Works the occurrence count out from the equivalent iCalendar recurrence rule.
        private static uint GetOccurrenceCount(DateTime start, DateTime endBy, string rrule)
        {
            var pattern = new CalendarRecurrence($"DTSTART:{start:yyyyMMdd}\r\nRRULE:{rrule}");
            return (uint)pattern.GenerateOccurrences(start, endBy).Count;
        }
    }
}

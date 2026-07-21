// Demonstrates how to create a daily recurring MAPI task that ends after N occurrences.

using System;
using Aspose.Email.Mapi;
using Aspose.Email.Calendar.Recurrences;

namespace Aspose.Email.Examples.MAPI
{
    internal static class EndAfterNoccurrences
    {
        public static void Run()
        {
            var offset = TimeZoneInfo.Local.GetUtcOffset(DateTime.Now);
            var startDate = new DateTime(2015, 7, 16).Add(offset);
            var dueDate = new DateTime(2015, 7, 16).Add(offset);
            var endByDate = new DateTime(2015, 8, 1).Add(offset);

            var task = new MapiTask("This is test task", "Sample Body", startDate, dueDate);
            task.State = MapiTaskState.NotAssigned;

            var rec = new MapiCalendarDailyRecurrencePattern
            {
                PatternType = MapiCalendarRecurrencePatternType.Day,
                Period = 1,
                WeekStartDay = DayOfWeek.Sunday,
                EndType = MapiCalendarRecurrenceEndType.EndAfterNOccurrences,
                OccurrenceCount = GetOccurrenceCount(startDate, endByDate, "FREQ=DAILY")
            };

            // Outlook rejects a pattern with no occurrences at all.
            if (rec.OccurrenceCount == 0)
                rec.OccurrenceCount = 1;

            task.Recurrence = rec;

            var outputPath = Data.Out/"Daily_out.msg";
            task.Save(outputPath, TaskSaveFormat.Msg);

            Console.WriteLine($"Task recurs every day, {rec.OccurrenceCount} occurrence(s).");
            Console.WriteLine($"Saved to {outputPath}");
        }

        private static uint GetOccurrenceCount(DateTime start, DateTime endBy, string rrule)
        {
            var pattern = new CalendarRecurrence($"DTSTART:{start:yyyyMMdd}\r\nRRULE:{rrule}");
            var dates = pattern.GenerateOccurrences(start, endBy);
            return (uint)dates.Count;
        }
    }
}

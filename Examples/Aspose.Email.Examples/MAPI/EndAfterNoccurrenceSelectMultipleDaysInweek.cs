// Demonstrates how to create a weekly recurring MAPI task on multiple weekdays ending after N occurrences.

using System;
using Aspose.Email.Mapi;
using Aspose.Email.Calendar.Recurrences;

namespace Aspose.Email.Examples.MAPI
{
    internal static class EndAfterNoccurrenceSelectMultipleDaysInWeek
    {
        public static void Run()
        {
            var offset = TimeZoneInfo.Local.GetUtcOffset(DateTime.Now);
            var startDate = new DateTime(2015, 7, 16).Add(offset);
            var dueDate = new DateTime(2015, 7, 16).Add(offset);
            var endByDate = new DateTime(2015, 9, 1).Add(offset);

            var task = new MapiTask("This is test task", "Sample Body", startDate, dueDate);
            task.State = MapiTaskState.NotAssigned;

            var rec = new MapiCalendarWeeklyRecurrencePattern
            {
                EndType = MapiCalendarRecurrenceEndType.EndAfterNOccurrences,
                PatternType = MapiCalendarRecurrencePatternType.Week,
                Period = 1,
                WeekStartDay = DayOfWeek.Sunday,
                DayOfWeek = MapiCalendarDayOfWeek.Friday | MapiCalendarDayOfWeek.Monday,
                OccurrenceCount = GetOccurrenceCount(startDate, endByDate, "FREQ=WEEKLY;BYDAY=FR,MO")
            };

            // Outlook rejects a pattern with no occurrences at all.
            if (rec.OccurrenceCount == 0)
                rec.OccurrenceCount = 1;

            task.Recurrence = rec;

            var outputPath = Data.Out/"EndAfterNoccurrenceSelectMultipleDaysInweek_out.msg";
            task.Save(outputPath, TaskSaveFormat.Msg);

            // DayOfWeek is a flags enum, so several days can be combined into one pattern.
            Console.WriteLine($"Task recurs on {rec.DayOfWeek}, {rec.OccurrenceCount} occurrence(s).");
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

// Demonstrates how to give an Outlook task a recurrence on several weekdays that
// repeats every second week and stops after a number of occurrences.

using System;
using Aspose.Email.Calendar.Recurrences;
using Aspose.Email.Mapi;

namespace Aspose.Email.Examples.MAPI
{
    internal static class SetWeeklyRecurrenceMultipleDaysInWeekWithInterval
    {
        public static void Run()
        {
            // Outlook stores task dates in UTC, so shift the local dates by the local
            // offset to keep the task on the intended day.
            var offset = TimeZoneInfo.Local.GetUtcOffset(DateTime.Now);

            var startDate = new DateTime(2015, 7, 16).Add(offset);
            var dueDate = new DateTime(2015, 7, 16).Add(offset);
            var endByDate = new DateTime(2015, 9, 1).Add(offset);

            var task = new MapiTask("This is test task", "Sample Body", startDate, dueDate)
            {
                State = MapiTaskState.NotAssigned
            };

            // DayOfWeek is a flags enum, so several days can be combined; Period 2 makes
            // the pattern repeat every second week rather than every week.
            var recurrence = new MapiCalendarWeeklyRecurrencePattern
            {
                EndType = MapiCalendarRecurrenceEndType.EndAfterNOccurrences,
                PatternType = MapiCalendarRecurrencePatternType.Week,
                Period = 2,
                WeekStartDay = DayOfWeek.Sunday,
                DayOfWeek = MapiCalendarDayOfWeek.Friday | MapiCalendarDayOfWeek.Monday,
                OccurrenceCount = GetOccurrenceCount(startDate, endByDate, "FREQ=WEEKLY;BYDAY=FR,MO;INTERVAL=2")
            };

            // Outlook rejects a pattern with no occurrences at all.
            if (recurrence.OccurrenceCount == 0)
                recurrence.OccurrenceCount = 1;

            task.Recurrence = recurrence;

            var outputPath = Data.Out/"SetWeeklyRecurrenceMultipleDaysInWeekWithInterval_out.msg";
            task.Save(outputPath, TaskSaveFormat.Msg);

            Console.WriteLine($"Task recurs on {recurrence.DayOfWeek} every {recurrence.Period} weeks, " +
                              $"{recurrence.OccurrenceCount} occurrence(s).");
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

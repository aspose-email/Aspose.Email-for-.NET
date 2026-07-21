// Demonstrates how to create a MapiTask with a monthly recurrence that stops on a
// given date.

using System;
using Aspose.Email.Calendar.Recurrences;
using Aspose.Email.Mapi;

namespace Aspose.Email.Examples.MAPI
{
    internal static class MonthlyEndAfterDate
    {
        public static void Run()
        {
            // Outlook stores task dates in UTC, so shift the local dates by the local
            // offset to keep the task on the intended day.
            var offset = TimeZoneInfo.Local.GetUtcOffset(DateTime.Now);
            var startDate = new DateTime(2015, 7, 1).Add(offset);
            var dueDate = new DateTime(2015, 7, 1).Add(offset);
            var endByDate = new DateTime(2015, 12, 31).Add(offset);

            var task = new MapiTask("This is test task", "Sample Body", startDate, dueDate)
            {
                State = MapiTaskState.NotAssigned
            };

            // Day 15 of every month, until the end date. EndAfterDate is what makes
            // EndDate the stopping condition rather than the occurrence count.
            var recurrence = new MapiCalendarMonthlyRecurrencePattern
            {
                Day = 15,
                Period = 1,
                PatternType = MapiCalendarRecurrencePatternType.Month,
                EndType = MapiCalendarRecurrenceEndType.EndAfterDate,
                EndDate = endByDate,
                OccurrenceCount = GetOccurrenceCount(startDate, endByDate, "FREQ=MONTHLY;BYMONTHDAY=15;INTERVAL=1"),
                WeekStartDay = DayOfWeek.Monday
            };

            // Outlook rejects a pattern with no occurrences at all.
            if (recurrence.OccurrenceCount == 0)
                recurrence.OccurrenceCount = 1;

            task.Recurrence = recurrence;

            var outputPath = Data.Out/"Monthly_out.msg";
            task.Save(outputPath, TaskSaveFormat.Msg);

            Console.WriteLine($"Task recurs on day {recurrence.Day} of every month until {endByDate:d}.");
            Console.WriteLine($"That is {recurrence.OccurrenceCount} occurrence(s).");
            Console.WriteLine($"Saved to {outputPath}");
        }

        // EndAfterDate patterns still need an occurrence count, so work it out from the
        // equivalent iCalendar recurrence rule over the same date range.
        private static uint GetOccurrenceCount(DateTime start, DateTime endBy, string rrule)
        {
            var pattern = new CalendarRecurrence($"DTSTART:{start:yyyyMMdd}\r\nRRULE:{rrule}");
            return (uint)pattern.GenerateOccurrences(start, endBy).Count;
        }
    }
}

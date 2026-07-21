// Demonstrates two ways of limiting a monthly recurrence: an occurrence count worked
// out from a date range, and a fixed count.

using System;
using Aspose.Email.Calendar.Recurrences;
using Aspose.Email.Mapi;

namespace Aspose.Email.Examples.MAPI
{
    internal static class MonthlyEndAfterNoccurrences
    {
        public static void Run()
        {
            // Outlook stores task dates in UTC, so shift the local dates by the local
            // offset to keep the task on the intended day.
            var offset = TimeZoneInfo.Local.GetUtcOffset(DateTime.Now);
            var startDate = new DateTime(2015, 7, 16).Add(offset);
            var dueDate = new DateTime(2015, 7, 16).Add(offset);
            var endByDate = new DateTime(2015, 12, 31).Add(offset);

            var task = new MapiTask("This is test task", "Sample Body", startDate, dueDate)
            {
                State = MapiTaskState.NotAssigned
            };

            // The count comes from how many times day 15 falls within the date range.
            var computed = new MapiCalendarMonthlyRecurrencePattern
            {
                Day = 15,
                Period = 1,
                PatternType = MapiCalendarRecurrencePatternType.Month,
                EndType = MapiCalendarRecurrenceEndType.EndAfterNOccurrences,
                OccurrenceCount = GetOccurrenceCount(startDate, endByDate, "FREQ=MONTHLY;BYMONTHDAY=15;INTERVAL=1"),
                WeekStartDay = DayOfWeek.Monday
            };

            // Outlook rejects a pattern with no occurrences at all.
            if (computed.OccurrenceCount == 0)
                computed.OccurrenceCount = 1;

            task.Recurrence = computed;
            var computedPath = Data.Out/"MonthlyEndAfterNoccurrences_out.msg";
            task.Save(computedPath, TaskSaveFormat.Msg);

            // Or simply state the count outright, with no date range involved.
            var fixedCount = new MapiCalendarMonthlyRecurrencePattern
            {
                Day = 15,
                Period = 1,
                PatternType = MapiCalendarRecurrencePatternType.Month,
                EndType = MapiCalendarRecurrenceEndType.EndAfterNOccurrences,
                OccurrenceCount = 5,
                WeekStartDay = DayOfWeek.Monday
            };

            task.Recurrence = fixedCount;
            var fixedPath = Data.Out/"MonthlyFixedOccurrences_out.msg";
            task.Save(fixedPath, TaskSaveFormat.Msg);

            Console.WriteLine($"Computed from {startDate:d}-{endByDate:d}: {computed.OccurrenceCount} occurrence(s)");
            Console.WriteLine($"  {computedPath}");
            Console.WriteLine($"Fixed count:                    {fixedCount.OccurrenceCount} occurrence(s)");
            Console.WriteLine($"  {fixedPath}");
        }

        // Works the occurrence count out from the equivalent iCalendar recurrence rule.
        private static uint GetOccurrenceCount(DateTime start, DateTime endBy, string rrule)
        {
            var pattern = new CalendarRecurrence($"DTSTART:{start:yyyyMMdd}\r\nRRULE:{rrule}");
            return (uint)pattern.GenerateOccurrences(start, endBy).Count;
        }
    }
}

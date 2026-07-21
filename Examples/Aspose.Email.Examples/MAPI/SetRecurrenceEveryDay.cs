// Demonstrates how to give an Outlook task a daily recurrence that ends on a date,
// once repeating every day and once every second day.

using System;
using Aspose.Email.Calendar.Recurrences;
using Aspose.Email.Mapi;

namespace Aspose.Email.Examples.MAPI
{
    internal static class SetRecurrenceEveryDay
    {
        public static void Run()
        {
            var startDate = new DateTime(2015, 7, 16);
            var endByDate = new DateTime(2015, 8, 1);
            var dueDate = new DateTime(2015, 7, 16);

            var task = new MapiTask("This is test task", "Sample Body", startDate, dueDate)
            {
                State = MapiTaskState.NotAssigned
            };

            // Period is the gap between occurrences: every day.
            var everyDay = new MapiCalendarDailyRecurrencePattern
            {
                PatternType = MapiCalendarRecurrencePatternType.Day,
                Period = 1,
                EndType = MapiCalendarRecurrenceEndType.EndAfterDate,
                OccurrenceCount = GetOccurrenceCount(startDate, endByDate, "FREQ=DAILY;INTERVAL=1"),
                EndDate = endByDate
            };

            task.Recurrence = everyDay;
            var everyDayPath = Data.Out/"SetRecurrenceEveryDay_out.msg";
            task.Save(everyDayPath, TaskSaveFormat.Msg);

            // The same range, but only every second day - hence fewer occurrences.
            var everySecondDay = new MapiCalendarDailyRecurrencePattern
            {
                PatternType = MapiCalendarRecurrencePatternType.Day,
                Period = 2,
                EndType = MapiCalendarRecurrenceEndType.EndAfterDate,
                OccurrenceCount = GetOccurrenceCount(startDate, endByDate, "FREQ=DAILY;INTERVAL=2"),
                EndDate = endByDate
            };

            task.Recurrence = everySecondDay;
            var everySecondDayPath = Data.Out/"SetRecurrenceEveryDayInterval_out.msg";
            task.Save(everySecondDayPath, TaskSaveFormat.Msg);

            Console.WriteLine($"Every day       -> {everyDay.OccurrenceCount} occurrence(s): {everyDayPath}");
            Console.WriteLine($"Every second day-> {everySecondDay.OccurrenceCount} occurrence(s): {everySecondDayPath}");
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

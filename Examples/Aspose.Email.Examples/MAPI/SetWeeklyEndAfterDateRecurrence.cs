// Demonstrates how to give an Outlook task a weekly recurrence that stops on a
// given date, once on a single weekday and once on two weekdays every second week.

using System;
using Aspose.Email.Calendar.Recurrences;
using Aspose.Email.Mapi;

namespace Aspose.Email.Examples.MAPI
{
    internal static class SetWeeklyEndAfterDateRecurrence
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

            // Every Friday, until the end date.
            var everyFriday = new MapiCalendarWeeklyRecurrencePattern
            {
                EndType = MapiCalendarRecurrenceEndType.EndAfterDate,
                PatternType = MapiCalendarRecurrencePatternType.Week,
                Period = 1,
                WeekStartDay = DayOfWeek.Sunday,
                DayOfWeek = MapiCalendarDayOfWeek.Friday,
                EndDate = endByDate,
                OccurrenceCount = GetOccurrenceCount(startDate, endByDate, "FREQ=WEEKLY;BYDAY=FR;INTERVAL=1")
            };

            // Outlook rejects a pattern with no occurrences at all.
            if (everyFriday.OccurrenceCount == 0)
                everyFriday.OccurrenceCount = 1;

            task.Recurrence = everyFriday;
            var singleDayPath = Data.Out/"SetWeeklyEndAfterDateEveryDayRecurrence_out.msg";
            task.Save(singleDayPath, TaskSaveFormat.Msg);

            // DayOfWeek is a flags enum, so several days can be combined - here Monday
            // and Friday, and Period 2 makes it every second week.
            var twoDaysEverySecondWeek = new MapiCalendarWeeklyRecurrencePattern
            {
                EndType = MapiCalendarRecurrenceEndType.EndAfterDate,
                PatternType = MapiCalendarRecurrencePatternType.Week,
                Period = 2,
                WeekStartDay = DayOfWeek.Sunday,
                EndDate = endByDate,
                DayOfWeek = MapiCalendarDayOfWeek.Friday | MapiCalendarDayOfWeek.Monday,
                OccurrenceCount = GetOccurrenceCount(startDate, endByDate, "FREQ=WEEKLY;BYDAY=FR,MO;INTERVAL=2")
            };

            task.Recurrence = twoDaysEverySecondWeek;
            var multipleDaysPath = Data.Out/"SetWeeklyEndAfterDateMultipleDaysRecurrence_out.msg";
            task.Save(multipleDaysPath, TaskSaveFormat.Msg);

            Console.WriteLine($"Every Friday             -> {everyFriday.OccurrenceCount} occurrence(s): {singleDayPath}");
            Console.WriteLine($"Mon+Fri every 2nd week   -> {twoDaysEverySecondWeek.OccurrenceCount} occurrence(s): {multipleDaysPath}");
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

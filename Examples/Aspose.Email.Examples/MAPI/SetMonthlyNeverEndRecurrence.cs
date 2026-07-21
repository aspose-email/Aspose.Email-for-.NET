// Demonstrates how to give an Outlook task a monthly recurrence that never ends.

using System;
using Aspose.Email.Mapi;

namespace Aspose.Email.Examples.MAPI
{
    internal static class SetMonthlyNeverEndRecurrence
    {
        public static void Run()
        {
            // Outlook stores task dates in UTC, so shift the local dates by the local
            // offset to keep the task on the intended day.
            var offset = TimeZoneInfo.Local.GetUtcOffset(DateTime.Now);

            var startDate = new DateTime(2015, 7, 1).Add(offset);
            var dueDate = new DateTime(2015, 7, 1).Add(offset);

            var task = new MapiTask("This is test task", "Sample Body", startDate, dueDate)
            {
                State = MapiTaskState.NotAssigned
            };

            // Day 15 of every month, for ever.
            var recurrence = new MapiCalendarMonthlyRecurrencePattern
            {
                Day = 15,
                Period = 1,
                PatternType = MapiCalendarRecurrencePatternType.Month,
                EndType = MapiCalendarRecurrenceEndType.NeverEnd,
                WeekStartDay = DayOfWeek.Monday
            };

            task.Recurrence = recurrence;

            var outputPath = Data.Out/"SetMonthlyNeverEndRecurrence_out.msg";
            task.Save(outputPath, TaskSaveFormat.Msg);

            Console.WriteLine($"Task recurs on day {recurrence.Day} of every month and never ends.");
            Console.WriteLine($"Saved to {outputPath}");
        }
    }
}

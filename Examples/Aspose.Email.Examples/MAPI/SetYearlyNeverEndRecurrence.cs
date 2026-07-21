// Demonstrates how to give an Outlook task a yearly recurrence that never ends.

using System;
using Aspose.Email.Mapi;

namespace Aspose.Email.Examples.MAPI
{
    internal static class SetYearlyNeverEndRecurrence
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

            // A yearly pattern is a monthly one with a period of 12 months.
            var recurrence = new MapiCalendarMonthlyRecurrencePattern
            {
                Day = 15,
                Period = 12,
                PatternType = MapiCalendarRecurrencePatternType.Month,
                EndType = MapiCalendarRecurrenceEndType.NeverEnd
            };

            // Outlook rejects a pattern with no occurrences at all.
            if (recurrence.OccurrenceCount == 0)
                recurrence.OccurrenceCount = 1;

            task.Recurrence = recurrence;

            var outputPath = Data.Out/"SetYearlyNeverEndRecurrence_out.msg";
            task.Save(outputPath, TaskSaveFormat.Msg);

            Console.WriteLine($"Task recurs on day {recurrence.Day} every {recurrence.Period} months and never ends.");
            Console.WriteLine($"Saved to {outputPath}");
        }
    }
}

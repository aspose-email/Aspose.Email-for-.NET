// Demonstrates how to set daily, weekly, monthly, and yearly recurrence patterns on a
// MAPI task, saving one file per pattern.

using System;
using Aspose.Email.Mapi;

namespace Aspose.Email.Examples.MAPI
{
    internal static class AddRecurrenceToMapiTask
    {
        public static void Run()
        {
            var startDate = new DateTime(2015, 04, 30, 10, 00, 00);
            var task = new MapiTask("abc", "def", startDate, startDate.AddHours(1))
            {
                State = MapiTaskState.NotAssigned
            };

            // NeverEnd patterns need no occurrence count, hence the zeros below.
            Save(task, "AsposeDaily_out.msg", "every day", new MapiCalendarDailyRecurrencePattern
            {
                PatternType = MapiCalendarRecurrencePatternType.Day,
                Period = 1,
                WeekStartDay = DayOfWeek.Sunday,
                EndType = MapiCalendarRecurrenceEndType.NeverEnd,
                OccurrenceCount = 0
            });

            Save(task, "AsposeWeekly_out.msg", "every Wednesday", new MapiCalendarWeeklyRecurrencePattern
            {
                PatternType = MapiCalendarRecurrencePatternType.Week,
                Period = 1,
                DayOfWeek = MapiCalendarDayOfWeek.Wednesday,
                EndType = MapiCalendarRecurrenceEndType.NeverEnd,
                OccurrenceCount = 0
            });

            Save(task, "AsposeMonthly_out.msg", "day 30 of every month", new MapiCalendarMonthlyRecurrencePattern
            {
                PatternType = MapiCalendarRecurrencePatternType.Month,
                Period = 1,
                EndType = MapiCalendarRecurrenceEndType.NeverEnd,
                Day = 30,
                OccurrenceCount = 0,
                WeekStartDay = DayOfWeek.Sunday
            });

            // A yearly pattern is a monthly one with a period of 12 months.
            Save(task, "AsposeYearly_out.msg", "every 12 months", new MapiCalendarMonthlyRecurrencePattern
            {
                PatternType = MapiCalendarRecurrencePatternType.Month,
                EndType = MapiCalendarRecurrenceEndType.NeverEnd,
                OccurrenceCount = 10,
                Period = 12
            });
        }

        private static void Save(MapiTask task, string fileName, string description, MapiCalendarRecurrencePattern pattern)
        {
            task.Recurrence = pattern;

            var path = Data.Out/fileName;
            task.Save(path, TaskSaveFormat.Msg);

            Console.WriteLine($"{description,-22} -> {path}");
        }
    }
}

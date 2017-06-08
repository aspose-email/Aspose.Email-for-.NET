using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Aspose.Email;
using Aspose.Email.Examples.CSharp.Email;
using Aspose.Email.Mime;
using Aspose.Email.Mapi;
using Aspose.Email.Storage.Pst;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/

namespace Aspose.Email.Examples.CSharp.Email.Outlook
{
    class AddRecurrenceToMapiTask
    {
        public static void Run()
        {
            // The path to the File directory.
            string dataDir = RunExamples.GetDataDir_Outlook();

            // ExStart:AddRecurrenceToMapiTask
            DateTime startDate = new DateTime(2015, 04, 30, 10, 00, 00);
            MapiTask task = new MapiTask("abc", "def", startDate, startDate.AddHours(1));
            task.State = MapiTaskState.NotAssigned;

            // Set the weekly recurrence
            var rec = new MapiCalendarDailyRecurrencePattern
            {
                PatternType = MapiCalendarRecurrencePatternType.Day,
                Period = 1,
                WeekStartDay = DayOfWeek.Sunday,
                EndType = MapiCalendarRecurrenceEndType.NeverEnd,
                OccurrenceCount = 0,
            };
            task.Recurrence = rec;
            task.Save(dataDir + "AsposeDaily_out.msg", TaskSaveFormat.Msg);

            // Set the weekly recurrence
            var rec1 = new MapiCalendarWeeklyRecurrencePattern
            {
                PatternType = MapiCalendarRecurrencePatternType.Week,
                Period = 1,
                DayOfWeek = MapiCalendarDayOfWeek.Wednesday,
                EndType = MapiCalendarRecurrenceEndType.NeverEnd,
                OccurrenceCount = 0,
            };
            task.Recurrence = rec1;
            task.Save(dataDir + "AsposeWeekly_out.msg", TaskSaveFormat.Msg);

            // Set the monthly recurrence
            var recMonthly = new MapiCalendarMonthlyRecurrencePattern
            {
                PatternType = MapiCalendarRecurrencePatternType.Month,
                Period = 1,
                EndType = MapiCalendarRecurrenceEndType.NeverEnd,
                Day = 30,
                OccurrenceCount = 0,
                WeekStartDay = DayOfWeek.Sunday,
            };
            task.Recurrence = recMonthly;
            //task.Save(dataDir + "AsposeMonthly_out.msg", TaskSaveFormat.Msg);

            // Set the yearly recurrence
            var recYearly = new MapiCalendarMonthlyRecurrencePattern
            {
                PatternType = MapiCalendarRecurrencePatternType.Month,
                EndType = MapiCalendarRecurrenceEndType.NeverEnd,
                OccurrenceCount = 10,
                Period = 12,
            };
            task.Recurrence = recYearly;
            //task.Save(dataDir + "AsposeYearly_out.msg", TaskSaveFormat.Msg);
            // ExEnd:AddRecurrenceToMapiTask
        }
    }
}
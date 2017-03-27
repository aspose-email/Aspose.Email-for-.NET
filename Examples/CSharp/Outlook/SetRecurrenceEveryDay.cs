using Aspose.Email.Mapi;
using Aspose.Email.Calendar.Recurrences;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

/* This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET 
   API reference when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq 
   for more information. If you do not wish to use NuGet, you can manually download 
   Aspose.Email for .NET API from http://www.aspose.com/downloads, 
   install it and then add its reference to this project. For any issues, questions or suggestions 
   please feel free to contact us using http://www.aspose.com/community/forums/default.aspx            
*/

namespace Aspose.Email.Examples.CSharp.Email.Outlook
{
    class SetRecurrenceEveryDay
    {
        public static void Run()
        {
            string dataDir = RunExamples.GetDataDir_Outlook();

            DateTime StartDate = new DateTime(2015, 7, 16);
            DateTime endByDate = new DateTime(2015, 8, 1);
            DateTime DueDate = new DateTime(2015, 7, 16);
            MapiTask task = new MapiTask("This is test task", "Sample Body", StartDate, DueDate);
            task.State = MapiTaskState.NotAssigned;

            // ExStart:SetRecurrenceEveryDay
            // Set the Daily recurrence
            var record = new MapiCalendarDailyRecurrencePattern
            {
                PatternType = MapiCalendarRecurrencePatternType.Day,
                Period = 1,
                EndType = MapiCalendarRecurrenceEndType.EndAfterDate,
                OccurrenceCount = GetOccurrenceCount(StartDate, endByDate, "FREQ=DAILY;INTERVAL=1"),
                EndDate = endByDate
            };
            // ExEnd:SetRecurrenceEveryDay

            task.Recurrence = record;
            task.Save(dataDir + "SetRecurrenceEveryDay_out.msg", TaskSaveFormat.Msg);

            // ExStart:SetRecurrenceEveryDayInterval
            // Set the Daily recurrence
            var record1 = new MapiCalendarDailyRecurrencePattern
            {
                PatternType = MapiCalendarRecurrencePatternType.Day,
                Period = 2,
                EndType = MapiCalendarRecurrenceEndType.EndAfterDate,
                OccurrenceCount = GetOccurrenceCount(StartDate, endByDate, "FREQ=DAILY;INTERVAL=2"),
                EndDate = endByDate
            };
            // ExEnd:SetRecurrenceEveryDayInterval

            task.Recurrence = record;
            task.Save(dataDir + "SetRecurrenceEveryDayInterval_out.msg", TaskSaveFormat.Msg);

        }

        private static uint GetOccurrenceCount(DateTime start, DateTime endBy, string rrule)
        {
            CalendarRecurrence pattern = new CalendarRecurrence(string.Format("DTSTART:{0}\r\nRRULE:{1}", start.ToString("yyyyMMdd"), rrule));
            DateCollection dates = pattern.GenerateOccurrences(start, endBy);
            return (uint)dates.Count;
        }

    }
}

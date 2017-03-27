using System;
using Aspose.Email.Mapi;
using Aspose.Email.Calendar.Recurrences;

/* This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET 
   API reference when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq 
   for more information. If you do not wish to use NuGet, you can manually download 
   Aspose.Email for .NET API from http://www.aspose.com/downloads, 
   install it and then add its reference to this project. For any issues, questions or suggestions 
   please feel free to contact us using http://www.aspose.com/community/forums/default.aspx            
*/

namespace Aspose.Email.Examples.CSharp.Email.Outlook
{
    class EndAfterNoccurrences
    {
        public static void Run()
        {
            // ExStart:EndAfterNoccurrences
            // The path to the File directory.
            string dataDir = RunExamples.GetDataDir_Outlook();

            TimeZone localZone = TimeZone.CurrentTimeZone;
            TimeSpan timeSpan = localZone.GetUtcOffset(DateTime.Now);
            DateTime StartDate = new DateTime(2015, 7, 16);
            StartDate = StartDate.Add(timeSpan);

            DateTime DueDate = new DateTime(2015, 7, 16);
            DateTime endByDate = new DateTime(2015, 8, 1);
            DueDate = DueDate.Add(timeSpan);
            endByDate = endByDate.Add(timeSpan);
 
            MapiTask task = new MapiTask("This is test task", "Sample Body", StartDate, DueDate);
            task.State = MapiTaskState.NotAssigned;

            // Set the Daily recurrence
            var rec = new MapiCalendarDailyRecurrencePattern
            {
                PatternType = MapiCalendarRecurrencePatternType.Day,
                Period = 1,
                WeekStartDay = DayOfWeek.Sunday,
                EndType = MapiCalendarRecurrenceEndType.EndAfterNOccurrences,
                OccurrenceCount = GetOccurrenceCount(StartDate, endByDate, "FREQ=DAILY"),
            };

            if (rec.OccurrenceCount==0)
            {
                rec.OccurrenceCount = 1;
            }

            task.Recurrence = rec;
            task.Save(dataDir + "Daily_out.msg", TaskSaveFormat.Msg);
            // ExEnd:EndAfterNoccurrences
        }
        private static uint GetOccurrenceCount(DateTime start, DateTime endBy, string rrule)
        {
            // ExStart:GetOccurrenceCount
            CalendarRecurrence pattern = new CalendarRecurrence(string.Format("DTSTART:{0}\r\nRRULE:{1}", start.ToString("yyyyMMdd"), rrule));
            DateCollection dates = pattern.GenerateOccurrences(start, endBy);       
            return (uint)dates.Count;
            // ExEnd:GetOccurrenceCount
        }
    }
}

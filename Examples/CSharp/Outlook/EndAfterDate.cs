using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Aspose.Email.Storage.Pst;
using Aspose.Email;
using Aspose.Email.Mapi;

/* This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET 
   API reference when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq 
   for more information. If you do not wish to use NuGet, you can manually download 
   Aspose.Email for .NET API from http://www.aspose.com/downloads, 
   install it and then add its reference to this project. For any issues, questions or suggestions 
   please feel free to contact us using http://www.aspose.com/community/forums/default.aspx            
*/

namespace Aspose.Email.Examples.CSharp.Email.Outlook
{
    class EndAfterDate
    {
        public static void Run()
        {
            // The path to the File directory.
            string dataDir = RunExamples.GetDataDir_Outlook();
            TimeZone localZone = TimeZone.CurrentTimeZone;
            TimeSpan timeSpan = localZone.GetUtcOffset(DateTime.Now);
            DateTime StartDate = new DateTime(2015, 7, 1);
            StartDate = StartDate.Add(timeSpan);

            DateTime DueDate = new DateTime(2015, 7, 1);
            DateTime endByDate = new DateTime(2020, 12, 31);
            DueDate = DueDate.Add(timeSpan);
            endByDate = endByDate.Add(timeSpan);

            MapiTask task = new MapiTask("This is test task", "Sample Body", StartDate, DueDate);
            task.State = MapiTaskState.NotAssigned;

            // Set the Yearly recurrence
            var rec = new MapiCalendarMonthlyRecurrencePattern
            {
                Day = 15,
                Period = 12,
                PatternType = MapiCalendarRecurrencePatternType.Day,
                EndType = MapiCalendarRecurrenceEndType.None,
                OccurrenceCount = 3,
            };

            task.Recurrence = rec;
            task.Save(dataDir + "EndAfterDate_out.msg", TaskSaveFormat.Msg);
        }

    }
}

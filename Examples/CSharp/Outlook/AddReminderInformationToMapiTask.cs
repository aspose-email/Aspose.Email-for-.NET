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
    class AddReminderInformationToMapiTask
    {
        public static void Run()
        {
            // ExStart:AddReminderInformationToMapiTask
            // The path to the File directory.
            string dataDir = RunExamples.GetDataDir_Outlook();

            // Create MapiTask and set Task Properties
            MapiTask testTask = new MapiTask("task with reminder", "this is a body", DateTime.Now, DateTime.Now.AddHours(1));
            testTask.ReminderSet = true;
            testTask.ReminderTime = DateTime.Now;
            testTask.ReminderFileParameter =dataDir + "Alarm01.wav";
            testTask.Save(dataDir + "AddReminderInformationToMapiTask_out", TaskSaveFormat.Msg);
            // ExEnd:AddReminderInformationToMapiTask
        }
    }
}
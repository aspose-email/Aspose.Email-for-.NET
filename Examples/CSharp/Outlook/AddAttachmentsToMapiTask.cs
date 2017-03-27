using System;
using System.Text;
using Aspose.Email.Mapi;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/

namespace Aspose.Email.Examples.CSharp.Email.Outlook
{
    class AddAttachmentsToMapiTask
    {
        public static void Run()
        {
            // ExStart:AddAttachmentsToMapiTask
            // The path to the File directory.
            string dataDir = RunExamples.GetDataDir_Outlook();

            string attachmentContent = "Test attachment body";
            string attachmentName = "Test attachment name";

            MapiTask testTask = new MapiTask("Task with attacment", "Test body of task with attacment", DateTime.Now, DateTime.Now.AddHours(1));
            testTask.Attachments.Add(attachmentName, Encoding.Unicode.GetBytes(attachmentContent));
            testTask.Save(dataDir + "AddAttachmentsToMapiTask_out", TaskSaveFormat.Msg);
            // ExEnd:AddAttachmentsToMapiTask
        }
    }
}
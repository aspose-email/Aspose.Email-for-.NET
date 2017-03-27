using System;
using System.IO;
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
    class AddAttachmentsToMapiJournal
    {
        public static void Run()
        {
            // The path to the File directory.
            // ExStart:AddAttachmentsToMapiJournal
            string dataDir = RunExamples.GetDataDir_Outlook();

            string[] attachFileNames = new string[] { dataDir + "Desert.jpg", dataDir + "download.png" };

            MapiJournal journal = new MapiJournal("testJournal", "This is a test journal", "Phone call", "Phone call");
            journal.StartTime = DateTime.Now;
            journal.EndTime = journal.StartTime.AddHours(1);
            journal.Companies = new string[] { "company 1", "company 2", "company 3" };

            foreach (string attach in attachFileNames)
            {
                journal.Attachments.Add(attach,File.ReadAllBytes(attach));
            }
            journal.Save(dataDir + "AddAttachmentsToMapiJournal_out.msg");
            // ExEnd:AddAttachmentsToMapiJournal
        }
    }
}
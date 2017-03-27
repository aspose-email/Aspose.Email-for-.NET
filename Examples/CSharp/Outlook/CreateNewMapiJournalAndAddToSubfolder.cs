using System;
using Aspose.Email.Mapi;
using Aspose.Email.Storage.Pst;
using System.IO;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/

namespace Aspose.Email.Examples.CSharp.Email.Outlook
{
    class CreateNewMapiJournalAndAddToSubfolder
    {
        public static void Run()
        {
            // The path to the File directory.
            // ExStart:CreateNewMapiJournalAndAddToSubfolder
            string dataDir = RunExamples.GetDataDir_Outlook();
            MapiJournal journal = new MapiJournal("daily record", "called out in the dark", "Phone call", "Phone call");
            journal.StartTime = DateTime.Now;
            journal.EndTime = journal.StartTime.AddHours(1);

            string path = dataDir + "CreateNewMapiJournalAndAddToSubfolder_out.pst";

            if (File.Exists(path))
            {
                File.Delete(path);
            }

            using (PersonalStorage personalStorage = PersonalStorage.Create(dataDir + "CreateNewMapiJournalAndAddToSubfolder_out.pst", FileFormatVersion.Unicode))
            {
                FolderInfo journalFolder = personalStorage.CreatePredefinedFolder("Journal", StandardIpmFolder.Journal);
                journalFolder.AddMapiMessageItem(journal);
            }
            // ExEnd:CreateNewMapiJournalAndAddToSubfolder
        }
    }
}
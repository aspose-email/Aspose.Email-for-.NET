using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Aspose.Email;
using Aspose.Email.Examples.CSharp.Email;
using Aspose.Email.Mime;
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
    class CreateNewMapiCalendarAndAddToCalendarSubfolder
    {
        public static void Run()
        {
            // The path to the File directory.
            // ExStart:CreateNewMapiCalendarAndAddToCalendarSubfolder
            string dataDir = RunExamples.GetDataDir_Outlook();
            MapiMessage message = MapiMessage.FromFile(dataDir + "Note.msg");

            // Note #1
            MapiNote note1 = (MapiNote)message.ToMapiMessageItem();
            note1.Subject = "Yellow color note";
            note1.Body = "This is a yellow color note";

            // Note #2
            MapiNote note2 = (MapiNote)message.ToMapiMessageItem();
            note2.Subject = "Pink color note";
            note2.Body = "This is a pink color note";
            note2.Color = NoteColor.Pink;

            // Note #3
            MapiNote note3 = (MapiNote)message.ToMapiMessageItem();
            note2.Subject = "Blue color note";
            note2.Body = "This is a blue color note";
            note2.Color = NoteColor.Blue;
            note3.Height = 500;
            note3.Width = 500;

            if (File.Exists(dataDir + "SampleNote_out.pst"))
            {
                File.Delete(dataDir + "SampleNote_out.pst");
            }

            using (PersonalStorage pst = PersonalStorage.Create(dataDir + "SampleNote_out.pst", FileFormatVersion.Unicode))
            {
                FolderInfo notesFolder = pst.CreatePredefinedFolder("Notes", StandardIpmFolder.Notes);
                notesFolder.AddMapiMessageItem(note1);
                notesFolder.AddMapiMessageItem(note2);
                notesFolder.AddMapiMessageItem(note3);
            }
            // ExEnd:CreateNewMapiCalendarAndAddToCalendarSubfolder
        }
    }
}
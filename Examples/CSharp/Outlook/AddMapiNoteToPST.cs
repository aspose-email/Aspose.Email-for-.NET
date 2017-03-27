using System.IO;
using Aspose.Email.Mapi;
using Aspose.Email.Storage.Pst;

/* This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET 
   API reference when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq 
   for more information. If you do not wish to use NuGet, you can manually download 
   Aspose.Email for .NET API from http://www.aspose.com/downloads, 
   install it and then add its reference to this project. For any issues, questions or suggestions 
   please feel free to contact us using http://www.aspose.com/community/forums/default.aspx            
*/

namespace Aspose.Email.Examples.CSharp.Email.Outlook
{
    class AddMapiNoteToPST
    {
        public static void Run()
        {
            // ExStart:AddMapiNoteToPST
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Outlook();

            MapiMessage mess = MapiMessage.FromFile(dataDir + "Note.msg");

            // Create three Notes  
            MapiNote note1 = (MapiNote)mess.ToMapiMessageItem();
            note1.Subject = "Yellow color note";
            note1.Body = "This is a yellow color note";

            MapiNote note2 = (MapiNote)mess.ToMapiMessageItem();
            note2.Subject = "Pink color note";
            note2.Body = "This is a pink color note";
            note2.Color = NoteColor.Pink;

            MapiNote note3 = (MapiNote)mess.ToMapiMessageItem();
            note2.Subject = "Blue color note";
            note2.Body = "This is a blue color note";
            note2.Color = NoteColor.Blue;
            note3.Height = 500;
            note3.Width = 500;

            string path = dataDir + "AddMapiNoteToPST_out.pst";

            if (File.Exists(path))
            {
                File.Delete(path);
            }

            using (PersonalStorage personalStorage = PersonalStorage.Create(dataDir + "AddMapiNoteToPST_out.pst", FileFormatVersion.Unicode))
            {
                FolderInfo notesFolder = personalStorage.CreatePredefinedFolder("Notes", StandardIpmFolder.Notes);
                notesFolder.AddMapiMessageItem(note1);
                notesFolder.AddMapiMessageItem(note2);
                notesFolder.AddMapiMessageItem(note3);
            }
            // ExEnd:AddMapiNoteToPST
        }
    }
}
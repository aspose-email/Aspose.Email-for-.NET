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
    class CreatAndSaveAnOutlookNote
    {
        public static void Run()
        {
            // ExStart:CreatAndSaveAnOutlookNote
            // The path to the File directory.
            string dataDir = RunExamples.GetDataDir_Outlook();
            // Create MapiNote and set Properties 
            MapiNote note3 = new MapiNote();
            note3.Subject = "Blue color note";
            note3.Body = "This is a blue color note";
            note3.Color = NoteColor.Blue;
            note3.Height = 500;
            note3.Width = 500;
            note3.Save(dataDir + "MapiNote_out.msg", NoteSaveFormat.Msg);
            // ExEnd:CreatAndSaveAnOutlookNote
        }
    }
}
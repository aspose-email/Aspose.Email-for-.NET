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
    class AddFilesToPST
    {
        public static void Run()
        {
            // The path to the File directory.            
            string dataDir = RunExamples.GetDataDir_Outlook();

            string path = dataDir + "Ps1_out.pst";

            if (File.Exists(path))
            {
                File.Delete(path);
            }

            // ExStart:AddFilesToPST
            using (PersonalStorage personalStorage = PersonalStorage.Create(dataDir + "Ps1_out.pst", FileFormatVersion.Unicode))
            {
                FolderInfo folder = personalStorage.RootFolder.AddSubFolder("Files");

                // Add Document.doc file with the "IPM.Document" message class by default.
                folder.AddFile(dataDir + "attachment_1.doc", null);
            }
            // ExEnd:AddFilesToPST
        }
    }
}
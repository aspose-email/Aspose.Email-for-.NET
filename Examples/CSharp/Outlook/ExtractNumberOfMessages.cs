using System.IO;
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
    class ExtractNumberOfMessages
    {
        public static void Run()
        {            
            // The path to the file directory.
            string dataDir = RunExamples.GetDataDir_Outlook();

            // Load the Outlook file
            string path = dataDir + "Sub.pst";

            // Save message to file
            using (PersonalStorage personalStorage = PersonalStorage.FromFile(path))
            {
                // ExStart:ExtractNnumberOfMessages
                FolderInfo inbox = personalStorage.RootFolder.GetSubFolder("Inbox");

                // Extracts messages starting from 10th index top and extract total 100 messages
                MessageInfoCollection messages = inbox.GetContents(10, 100);
                // ExEnd:ExtractNnumberOfMessages   
            }                    
        }
    }
}

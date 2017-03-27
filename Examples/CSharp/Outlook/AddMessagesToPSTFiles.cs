using Aspose.Email.Mapi;
using Aspose.Email.Storage.Pst;
using System.IO;

/* This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET 
   API reference when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq 
   for more information. If you do not wish to use NuGet, you can manually download 
   Aspose.Email for .NET API from http://www.aspose.com/downloads, 
   install it and then add its reference to this project. For any issues, questions or suggestions 
   please feel free to contact us using http://www.aspose.com/community/forums/default.aspx            
*/

namespace Aspose.Email.Examples.CSharp.Email.Outlook
{
    class AddMessagesToPSTFiles
    {
        public static void Run()
        {            
            // The path to the file directory.
            string dataDir = RunExamples.GetDataDir_Outlook() + "AddMessagesToPSTFiles_out.pst";

            if (File.Exists(dataDir))
            {
                File.Delete(dataDir);
            }
            else
            {
                
            }
            
            // ExStart:AddMessagesToPSTFiles
            // Create new PST            
            PersonalStorage personalStorage = PersonalStorage.Create(dataDir, FileFormatVersion.Unicode);

            // Add new folder "Inbox"
            personalStorage.RootFolder.AddSubFolder("Inbox");

            // Select the "Inbox" folder
            FolderInfo inboxFolder = personalStorage.RootFolder.GetSubFolder("Inbox");

            // Add some messages to "Inbox" folder
            inboxFolder.AddMessage(MapiMessage.FromFile(RunExamples.GetDataDir_Outlook() + "MapiMsgWithPoll.msg"));
            // ExEnd:AddMessagesToPSTFiles            
        }
    }
}

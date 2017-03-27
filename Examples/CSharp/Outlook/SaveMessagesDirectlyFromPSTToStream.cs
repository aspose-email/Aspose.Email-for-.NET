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
    class SaveMessagesDirectlyFromPSTToStream
    {
        public static void Run()
        {
            // ExStart:SaveMessagesDirectlyFromPSTToStream
            // The path to the file directory.
            string dataDir = RunExamples.GetDataDir_Outlook();

            // Load the Outlook file
            string path = dataDir + "PersonalStorage.pst";

            // Save message to MemoryStream
            using (PersonalStorage personalStorage = PersonalStorage.FromFile(path))
            {
                FolderInfo inbox = personalStorage.RootFolder.GetSubFolder("Inbox");
                foreach (MessageInfo messageInfo in inbox.EnumerateMessages())
                {
                    using (MemoryStream memeorystream = new MemoryStream())
                    {
                        personalStorage.SaveMessageToStream(messageInfo.EntryIdString, memeorystream);
                    }
                }
            }

            // Save message to file
            using (PersonalStorage pst = PersonalStorage.FromFile(path))
            {
                FolderInfo inbox = pst.RootFolder.GetSubFolder("Inbox");
                foreach (MessageInfo messageInfo in inbox.EnumerateMessages())
                {
                    using (FileStream fs = File.OpenWrite(messageInfo.Subject + ".msg"))
                    {
                        pst.SaveMessageToStream(messageInfo.EntryIdString, fs);
                    }
                }
            }

            using (PersonalStorage pst = PersonalStorage.FromFile(path))
            {
                FolderInfo inbox = pst.RootFolder.GetSubFolder("Inbox");
                
                // To enumerate entryId of messages you may use FolderInfo.EnumerateMessagesEntryId() method:
                foreach (string entryId in inbox.EnumerateMessagesEntryId())
                {
                    using (MemoryStream ms = new MemoryStream())
                    {
                        pst.SaveMessageToStream(entryId, ms);
                    }
                }
            }            
            // ExEnd:SaveMessagesDirectlyFromPSTToStream
        }
    }
}

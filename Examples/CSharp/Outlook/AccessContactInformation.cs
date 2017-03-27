using System;
using Aspose.Email;
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
    class AccessContactInformation
    {
        public static void Run()
        {
            // ExStart:AccessContactInformation
            // Load the Outlook file
            string dataDir = RunExamples.GetDataDir_Outlook();

            // Load the Outlook PST file
            PersonalStorage personalStorage = PersonalStorage.FromFile(dataDir + "SampleContacts.pst");
            // Get the Contacts folder
            FolderInfo folderInfo = personalStorage.RootFolder.GetSubFolder("Contacts");
            // Loop through all the contacts in this folder
            MessageInfoCollection messageInfoCollection = folderInfo.GetContents();
            foreach (MessageInfo messageInfo in messageInfoCollection)
            {
                // Get the contact information
                MapiMessage mapi = personalStorage.ExtractMessage(messageInfo);

                MapiContact contact = (MapiContact)mapi.ToMapiMessageItem();

                // Display some contents on screen
                Console.WriteLine("Name: " + contact.NameInfo.DisplayName);
                // Save to disk in MSG format
                if (contact.NameInfo.DisplayName != null)
                {
                    MapiMessage message = personalStorage.ExtractMessage(messageInfo);
                    // Get rid of illegal characters that cannot be used as a file name
                    string messageName = message.Subject.Replace(":", " ").Replace("\\", " ").Replace("?", " ").Replace("/", " ");
                    message.Save(dataDir + "Contacts\\" + messageName + "_out.msg");
                }
            }
            // ExEnd:AccessContactInformation
        }
    }
}
using System;
using Aspose.Email.Mapi;
using Aspose.Email.Storage.Pst;

/* This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET 
   API reference when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq 
   for more information. If you do not wish to use NuGet, you can manually download 
   Aspose.Email for .NET API from http://www.aspose.com/downloads, 
   install it and then add its reference to this project. For any issues, questions or suggestions 
   please feel free to contact us using http://www.aspose.com/community/forums/default.aspx            
 * 
*/
namespace Aspose.Email.Examples.CSharp.Email.Outlook
{
    class SaveContactInformation
    {
        public static void Run()
        {
            // ExStart:SaveContactInformation
            // Load the Outlook file
            string dataDir = RunExamples.GetDataDir_Outlook();
 
            // Load the Outlook PST file
            PersonalStorage personalStorage = PersonalStorage.FromFile(dataDir + "Outlook.pst");
            // Get the Contacts folder
            FolderInfo folderInfo = personalStorage.RootFolder.GetSubFolder("Contacts");
            // Loop through all the contacts in this folder
            MessageInfoCollection messageInfoCollection = folderInfo.GetContents();
            foreach (MessageInfo messageInfo in messageInfoCollection)
            {
                // Get the contact information
                MapiContact contact = (MapiContact)personalStorage.ExtractMessage(messageInfo).ToMapiMessageItem();
                // Display some contents on screen
                Console.WriteLine("Name: " + contact.NameInfo.DisplayName + " - " + messageInfo.EntryIdString);
                // Save to disk in vCard VCF format
                contact.Save(dataDir + "Contacts\\" + contact.NameInfo.DisplayName + ".vcf", ContactSaveFormat.VCard);
            }
            // ExEnd:SaveContactInformation
        }
    }
}
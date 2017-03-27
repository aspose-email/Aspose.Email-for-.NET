using System.IO;
using System;
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
    class SaveCalendarItems
    {
        public static void Run()
        {
            // ExStart:SaveCalendarItems
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Outlook();

            // Load the Outlook PST file
            PersonalStorage pst = PersonalStorage.FromFile(dataDir + "Sub.pst");
            // Get the Calendar folder
            FolderInfo folderInfo = pst.RootFolder.GetSubFolder("Inbox");
            // Loop through all the calendar items in this folder
            MessageInfoCollection messageInfoCollection = folderInfo.GetContents();
            foreach (MessageInfo messageInfo in messageInfoCollection)
            {
                // Get the calendar information
                MapiMessage calendar = (MapiMessage)pst.ExtractMessage(messageInfo).ToMapiMessageItem();
                // Display some contents on screen
                Console.WriteLine("Name: " + calendar.Subject);
                // Save to disk in ICS format
                calendar.Save(dataDir + @"\Calendar\" + calendar.Subject + "_out.ics");
            }
            // ExEnd:SaveCalendarItems
        }
    }
}
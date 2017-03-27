using System;
using Aspose.Email.Clients.Imap;
using Aspose.Email.Mime;

/* This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET 
   API reference when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq 
   for more information. If you do not wish to use NuGet, you can manually download 
   Aspose.Email for .NET API from http://www.aspose.com/downloads, 
   install it and then add its reference to this project. For any issues, questions or suggestions 
   please feel free to contact us using http://www.aspose.com/community/forums/default.aspx            
*/

namespace Aspose.Email.Examples.CSharp.Email.Outlook
{
    class CopyMultipleMessagesFromOneFoldertoAnother
    {
        public static void Run()
        {
            //ExStart:CopyMultipleMessagesFromOneFoldertoAnother
            using (ImapClient client = new ImapClient("exchange.aspose.com", "username", "password"))
            {
                // Create the destination folder
                string folderName = "EMAILNET-35242";
                if (!client.ExistFolder(folderName))
                    client.CreateFolder(folderName);
                try
                {
                    // Append a couple of messages to the server
                    MailMessage message1 = new MailMessage(
                        "asposeemail.test3@aspose.com",
                        "asposeemail.test3@aspose.com",
                        "EMAILNET-35242 - " + Guid.NewGuid(),
                        "EMAILNET-35242 Improvement of copy method.Add ability to 'copy' multiple messages per invocation.");
                    string uniqueId1 = client.AppendMessage(message1);

                    MailMessage message2 = new MailMessage(
                        "asposeemail.test3@aspose.com",
                        "asposeemail.test3@aspose.com",
                        "EMAILNET-35242 - " + Guid.NewGuid(),
                        "EMAILNET-35242 Improvement of copy method.Add ability to 'copy' multiple messages per invocation.");
                    string uniqueId2 = client.AppendMessage(message2);

                    // Verify that the messages have been added to the mailbox
                    client.SelectFolder(ImapFolderInfo.InBox);
                    ImapMessageInfoCollection msgsColl = client.ListMessages();
                    foreach (ImapMessageInfo msgInfo in msgsColl)
                        Console.WriteLine(msgInfo.Subject);

                    // Copy multiple messages using the CopyMessages command and Verify that messages are copied to the destination folder
                    client.CopyMessages(new[] { uniqueId1, uniqueId2 }, folderName, true);
                    client.SelectFolder(folderName);
                    msgsColl = client.ListMessages();
                    foreach (ImapMessageInfo msgInfo in msgsColl)
                        Console.WriteLine(msgInfo.Subject);
                }
                finally
                {
                    try
                    {
                        client.DeleteFolder(folderName);
                    }
                    catch { }
                }
            }
            //ExEnd:CopyMultipleMessagesFromOneFoldertoAnother
        }
    }
}


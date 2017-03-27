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

namespace Aspose.Email.Examples.CSharp.Email.IMAP
{
    class DeleteSingleMessage
    {
        public static void Run()
        {
            //ExStart:DeleteSingleMessage
            using (ImapClient client = new ImapClient("exchange.aspose.com", "username", "password"))
            {
                try
                {
                    Console.WriteLine(client.UidPlusSupported.ToString());
                    // Append some test messages
                    client.SelectFolder(ImapFolderInfo.InBox);
                    MailMessage message = new MailMessage("from@Aspose.com", "to@Aspose.com", "EMAILNET-35227 - " + Guid.NewGuid(), "EMAILNET-35227 Add ability in ImapClient to delete message");
                    string emailId = client.AppendMessage(message);

                    // Now verify that all the messages have been appended to the mailbox
                    ImapMessageInfoCollection messageInfoCol = null;
                    messageInfoCol = client.ListMessages();
                    Console.WriteLine(messageInfoCol.Count);

                    // Select the inbox folder and Delete message
                    client.SelectFolder(ImapFolderInfo.InBox);
                    client.DeleteMessage(emailId);
                    client.CommitDeletes();
                }
                finally
                {

                }
            }
            //ExEnd:DeleteSingleMessage
        }
    }
}


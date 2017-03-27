using System;
using System.Collections.Generic;
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
    class DeleteMultipleMessages
    {
        public static void Run()
        {
            //ExStart:DeleteMultipleMessages
            using (ImapClient client = new ImapClient("exchange.aspose.com", "username", "password"))
            {
                try
                {
                    Console.WriteLine(client.UidPlusSupported.ToString());
                    // Append some test messages
                    client.SelectFolder(ImapFolderInfo.InBox);
                    List<string> uidList = new List<string>();
                    const int messageNumber = 5;
                    for (int i = 0; i < messageNumber; i++)
                    {
                        MailMessage message = new MailMessage(
                            "from@Aspose.com",
                            "to@Aspose.com",
                            "EMAILNET-35226 - " + Guid.NewGuid(),
                            "EMAILNET-35226 Add ability in ImapClient to delete messages and change flags for set of messages");
                        string uid = client.AppendMessage(message);
                        uidList.Add(uid);
                    }

                    // Now verify that all the messages have been appended to the mailbox
                    ImapMessageInfoCollection messageInfoCol = null;
                    messageInfoCol = client.ListMessages();
                    Console.WriteLine(messageInfoCol.Count);

                    // Bulk Delete Messages and  Verify that the messages are deleted
                    client.DeleteMessages(uidList, true);
                    client.CommitDeletes(); 
                    messageInfoCol = null;
                    messageInfoCol = client.ListMessages();
                    Console.WriteLine(messageInfoCol.Count);
                }
                finally
                {

                }
            }
            //ExEnd:DeleteMultipleMessages
        }
    }
}


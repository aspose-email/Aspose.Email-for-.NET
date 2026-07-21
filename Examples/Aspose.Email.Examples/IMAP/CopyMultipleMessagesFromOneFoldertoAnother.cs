using System;
using Aspose.Email.Clients.Imap;
using Aspose.Email.Mime;

namespace Aspose.Email.Examples.IMAP
{
    class CopyMultipleMessagesFromOneFoldertoAnother
    {
        public static void Run()
        {
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
                    client.CopyMessages(new[] { uniqueId1, uniqueId2 }, folderName);

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
        }
    }
}


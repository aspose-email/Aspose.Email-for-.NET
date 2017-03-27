using System;
using Aspose.Email.Clients.Imap;
using Aspose.Email.Mime;

namespace Aspose.Email.Examples.CSharp.Email.IMAP
{
    class MoveMessage
    {
        static void Run()
        { 
            // ExStart: MoveMessage
            ///<summary>
            /// This example shows how to move a message from one folder of a mailbox to another one using the ImapClient API of Aspose.Email for .NET
            /// Available from Aspose.Email for .NET 6.4.0 onwards
            /// -------------- Available API Overload Members --------------
            /// Void ImapClient.MoveMessage(IConnection iConnection, int sequenceNumber, string folderName, bool commitDeletions)
            /// Void ImapClient.MoveMessage(IConnection iConnection, string uniqueId, string folderName, bool commitDeletions)
            /// Void ImapClient.MoveMessage(int sequenceNumber, string folderName, bool commitDeletions)
            /// Void ImapClient.MoveMessage(string uniqueId, string folderName, bool commitDeletions)
            /// Void ImapClient.MoveMessage(IConnection iConnection, int sequenceNumber, string folderName)
            /// Void ImapClient.MoveMessage(IConnection iConnection, string uniqueId, string folderName)
            /// Void ImapClient.MoveMessage(int sequenceNumber, string folderName)
            /// Void ImapClient.MoveMessage(string uniqueId, string folderName)
            ///</summary>

            using (ImapClient client = new ImapClient("host.domain.com", 993, "username", "password"))
            {
                string folderName = "EMAILNET-35151";
                if (!client.ExistFolder(folderName))
                    client.CreateFolder(folderName);
                try
                {
                    MailMessage message = new MailMessage(
                        "from@domain.com",
                        "to@domain.com",
                        "EMAILNET-35151 - " + Guid.NewGuid(),
                        "EMAILNET-35151 ImapClient: Provide option to Move Message");
                    client.SelectFolder(ImapFolderInfo.InBox);
                    // Append the new message to Inbox folder
                    string uniqueId = client.AppendMessage(ImapFolderInfo.InBox, message);
                    ImapMessageInfoCollection messageInfoCol1 = client.ListMessages();
                    Console.WriteLine(messageInfoCol1.Count);
                    // Now move the message to the folder EMAILNET-35151
                    client.MoveMessage(uniqueId, folderName);
                    client.CommitDeletes();
                    // Verify that the message was moved to the new folder
                    client.SelectFolder(folderName);
                    messageInfoCol1 = client.ListMessages();
                    Console.WriteLine(messageInfoCol1.Count);
                    // Verify that the message was moved from the Inbox
                    client.SelectFolder(ImapFolderInfo.InBox);
                    messageInfoCol1 = client.ListMessages();
                    Console.WriteLine(messageInfoCol1.Count);
                }
                finally
                {
                    try { client.DeleteFolder(folderName); }
                    catch { }
                }
            }
            // ExEnd: MoveMessage
        }
    }
}

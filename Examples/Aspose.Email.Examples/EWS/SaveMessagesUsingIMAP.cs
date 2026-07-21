using Aspose.Email.Clients;
using Aspose.Email.Clients.Imap;
using Aspose.Email.Mime;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Aspose.Email.Examples.EWS
{
    class SaveMessagesUsingIMAP
    {
        public static void Run()
        {
            try
            {
                ImapClient imapClient = new ImapClient("ex07sp1", "Administrator", "Evaluation1");
                imapClient.SecurityOptions = SecurityOptions.Auto;

                // Select the Inbox folder
                imapClient.SelectFolder(ImapFolderInfo.InBox);
                // Get the list of messages
                ImapMessageInfoCollection msgCollection = imapClient.ListMessages();
                foreach (ImapMessageInfo msgInfo in msgCollection)
                {
                    // Fetch the message from inbox using its SequenceNumber from msgInfo
                    MailMessage message = imapClient.FetchMessage(msgInfo.SequenceNumber);

                    // Save the message to disc now
                    message.Save(Data.Out + msgInfo.SequenceNumber + "_out.msg", SaveOptions.DefaultMsgUnicode);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
    }
}

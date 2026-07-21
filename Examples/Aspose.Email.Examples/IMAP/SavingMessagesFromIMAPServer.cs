using Aspose.Email.Clients.Imap;
using Aspose.Email.Mime;

namespace Aspose.Email.Examples.IMAP
{
    class SavingMessagesFromIMAPServer
    {
        public static void Run()
        {
            
            // Create an imapclient with host, user and password
            ImapClient client = new ImapClient("localhost", "user", "password");

            // Select the inbox folder and Get the message info collection
            client.SelectFolder(ImapFolderInfo.InBox);
            ImapMessageInfoCollection list = client.ListMessages();

            // Download each message
            for (int i = 0; i < list.Count; i++)
            {
                // Save the message in MSG format
                MailMessage message = client.FetchMessage(list[i].UniqueId);
                message.Save(Data.Out + list[i].UniqueId + "_out.msg", SaveOptions.DefaultMsgUnicode);
            }
        }
    }
}
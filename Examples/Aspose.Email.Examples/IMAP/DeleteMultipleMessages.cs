using System;
using System.Collections.Generic;
using Aspose.Email.Clients.Imap;
using Aspose.Email.Mime;

namespace Aspose.Email.Examples.IMAP
{
    class DeleteMultipleMessages
    {
        public static void Run()
        {
            using (var client = ClientBuilder.Imap(AuthType.ModernWithDelegatedPermission))
            {
                client.SelectFolder(ImapFolderInfo.InBox);

                // Append test messages
                var emlList = new List<MailMessage>();

                for (var i = 0; i < 5; i++)
                {
                    var eml = new MailMessage("from@from.com", "to@to.com")
                    {
                        Subject = $"Message to delete {i}",
                        Body = "Hey! This Message will be deleted!"
                    };

                    emlList.Add(eml);
                }

                var appendMessagesResult = client.AppendMessages(emlList);

                // Bulk Delete appended Messages
                client.DeleteMessages(appendMessagesResult.Succeeded.Values, true);
                client.CommitDeletes();
            }
        }
    }
}


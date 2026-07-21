using System;
using Aspose.Email.Clients;
using Aspose.Email.Clients.Imap;
using Aspose.Email.Mime;

namespace Aspose.Email.Examples.IMAP
{
    class DeleteSingleMessage
    {
        public static void Run()
        {
            using (var client = ClientBuilder.Imap(AuthType.ModernWithDelegatedPermission))
            {
                client.SecurityOptions = SecurityOptions.SSLImplicit;

                // Append some test message
                client.SelectFolder(ImapFolderInfo.InBox);

                var eml = new MailMessage("from@from.com", "to@to.com")
                {
                    Subject = "Message to delete",
                    Body = "Hey! This Message will be deleted!"
                };
                var emlId = client.AppendMessage(eml);

                var fetchedEml = client.FetchMessage(emlId);
                Console.WriteLine(fetchedEml.Subject);

                client.DeleteMessage(emlId);
                client.CommitDeletes();
            }
        }
    }
}


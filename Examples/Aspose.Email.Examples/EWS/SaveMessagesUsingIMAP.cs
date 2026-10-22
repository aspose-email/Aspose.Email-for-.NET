// Demonstrates saving messages from Exchange over IMAP as Outlook .msg files.
//
// FetchMessage downloads a message over IMAP into a MailMessage, and Save converts it
// to Unicode MSG. Only the five newest messages of the Inbox are saved, under Out.
// SaveMessagesInMSGFormatExchangeEWS does the same over EWS.

using System;
using System.Linq;
using Aspose.Email.Clients.Imap;

namespace Aspose.Email.Examples.EWS
{
    internal static class SaveMessagesUsingIMAP
    {
        public static void Run()
        {
            var outputDir = Data.OutSub("ExchangeImapMsg");

            using (var client = ClientBuilder.Imap(AuthType.ModernWithDelegatedPermission))
            {
                client.SelectFolder(ImapFolderInfo.InBox);

                var newest = client.ListMessages().OrderByDescending(info => info.InternalDate).Take(5).ToList();
                foreach (var info in newest)
                {
                    var message = client.FetchMessage(info.UniqueId);
                    message.Save(outputDir/(info.UniqueId + ".msg"), SaveOptions.DefaultMsgUnicode);
                    Console.WriteLine($"Saved {info.UniqueId}.msg  {message.Subject}");
                }

                Console.WriteLine($"\n{newest.Count} message(s) saved to {outputDir}");
            }
        }
    }
}

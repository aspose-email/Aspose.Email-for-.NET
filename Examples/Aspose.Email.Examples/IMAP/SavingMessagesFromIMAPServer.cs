// Demonstrates how to download messages and save them in Outlook's .msg format.
//
// FetchMessage parses the message into a MailMessage, and Save converts it to the format
// chosen with SaveOptions - here Unicode MSG, which Outlook opens directly.
// MessagesFromIMAPServerToDisk keeps the original .eml data instead.

using System;
using System.Linq;
using Aspose.Email.Clients.Imap;

namespace Aspose.Email.Examples.IMAP
{
    internal static class SavingMessagesFromIMAPServer
    {
        public static void Run()
        {
            var outputDir = Data.OutSub("ImapMsg");

            using (var client = ClientBuilder.Imap(AuthType.ModernWithDelegatedPermission))
            {
                client.SelectFolder(ImapFolderInfo.InBox);

                var newest = client.ListMessages()
                    .OrderByDescending(info => info.InternalDate)
                    .Take(5)
                    .ToList();

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

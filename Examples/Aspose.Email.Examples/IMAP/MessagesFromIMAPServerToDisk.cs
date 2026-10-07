// Demonstrates how to save messages from the server straight to .eml files.
//
// SaveMessage writes the raw message as the server stores it, without parsing it into a
// MailMessage first - the quickest way to archive mail. The files can be opened by any
// mail program or loaded later with MailMessage.Load.

using System;
using System.Linq;
using Aspose.Email.Clients.Imap;

namespace Aspose.Email.Examples.IMAP
{
    internal static class MessagesFromIMAPServerToDisk
    {
        public static void Run()
        {
            var outputDir = Data.OutSub("ImapMessages");

            using (var client = ClientBuilder.Imap(AuthType.ModernWithDelegatedPermission))
            {
                client.SelectFolder(ImapFolderInfo.InBox);

                var newest = client.ListMessages()
                    .OrderByDescending(info => info.InternalDate)
                    .Take(5)
                    .ToList();

                foreach (var info in newest)
                {
                    var outputPath = outputDir/(info.UniqueId + ".eml");
                    client.SaveMessage(info.UniqueId, outputPath);
                    Console.WriteLine($"Saved {info.UniqueId}.eml  {info.Subject}");
                }

                Console.WriteLine($"\n{newest.Count} message(s) saved to {outputDir}");
            }
        }
    }
}

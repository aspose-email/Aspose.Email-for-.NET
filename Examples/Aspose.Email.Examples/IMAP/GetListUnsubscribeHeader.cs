// Demonstrates how to read the List-Unsubscribe header without downloading messages.
//
// Newsletters and mailing lists put an unsubscribe link or address into the
// List-Unsubscribe header (RFC 2369). ImapMessageInfo carries it in the summary that
// ListMessages returns, which makes it cheap to collect for a whole folder.

using System;
using System.Linq;
using Aspose.Email.Clients.Imap;

namespace Aspose.Email.Examples.IMAP
{
    internal static class GetListUnsubscribeHeader
    {
        public static void Run()
        {
            using (var client = ClientBuilder.Imap(AuthType.ModernWithDelegatedPermission))
            {
                client.SelectFolder(ImapFolderInfo.InBox);

                var withHeader = client.ListMessages()
                    .Where(info => !string.IsNullOrEmpty(info.ListUnsubscribe))
                    .ToList();

                Console.WriteLine($"{withHeader.Count} message(s) in the Inbox carry List-Unsubscribe:");

                foreach (var info in withHeader.Take(10))
                {
                    Console.WriteLine($"  {info.From}: {info.Subject}");
                    Console.WriteLine($"    {info.ListUnsubscribe}");
                }
            }
        }
    }
}

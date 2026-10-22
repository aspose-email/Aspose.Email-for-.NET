// Demonstrates that Exchange can also be reached over IMAP.
//
// Besides EWS, Exchange Online and Exchange Server offer IMAP, which suits code that
// already works with other IMAP servers. The "Imap" section of clientsettings.json
// points at outlook.office365.com, and the IMAP folder holds the full set of IMAP
// examples. EWS gives access to more of Exchange - calendars, contacts, tasks, rules.

using System;
using System.Linq;
using Aspose.Email.Clients.Imap;

namespace Aspose.Email.Examples.EWS
{
    internal static class ConnectExchangeServerUsingIMAP
    {
        public static void Run()
        {
            using (var client = ClientBuilder.Imap(AuthType.ModernWithDelegatedPermission))
            {
                client.SelectFolder(ImapFolderInfo.InBox);
                Console.WriteLine($"Connected to {client.Host} over IMAP; " +
                                  $"{client.CurrentFolder.TotalMessageCount} message(s) in the Inbox.");

                foreach (var info in client.ListMessages().OrderByDescending(i => i.InternalDate).Take(5))
                    Console.WriteLine($"  {info.InternalDate:g}  {info.Subject}");
            }
        }
    }
}

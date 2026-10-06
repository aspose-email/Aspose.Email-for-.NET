// Demonstrates searching on any header field, and passing server-specific search syntax
// through unchanged.
//
// HasHeader matches messages that have the header with the given text in its value; an
// empty value matches every message that has the header at all. CustomSearch adds raw
// search syntax the builder has no method for - here Gmail's X-GM-RAW, which accepts
// the same expressions as the Gmail search box.

using System;
using Aspose.Email.Clients.Imap;

namespace Aspose.Email.Examples.IMAP
{
    internal static class SearchByHeaderAndRawCriteria
    {
        public static void Run()
        {
            using (var client = ClientBuilder.Imap(AuthType.ModernWithDelegatedPermission))
            {
                client.SelectFolder(ImapFolderInfo.InBox);

                // Mailing-list traffic: anything that carries a List-Id header.
                var builder = new ImapQueryBuilder();
                builder.HasHeader("List-Id", "");
                Console.WriteLine($"Mailing-list messages:   {client.ListMessages(builder.GetQuery()).Count}");

                // Messages the sender marked as high priority.
                builder = new ImapQueryBuilder();
                builder.HasHeader("X-Priority", "1");
                Console.WriteLine($"High-priority messages:  {client.ListMessages(builder.GetQuery()).Count}");

                if (!client.GmExt1Supported)
                {
                    Console.WriteLine("\nX-GM-RAW needs Gmail; skipping the raw search.");
                    return;
                }

                builder = new ImapQueryBuilder();
                builder.CustomSearch("X-GM-RAW \"has:attachment older_than:1y\"");
                var old = client.ListMessages(builder.GetQuery());

                Console.WriteLine($"\nWith attachments, older than a year: {old.Count}");
                foreach (var info in old)
                    Console.WriteLine($"  {info.Date:d}  {info.Subject}");
            }
        }
    }
}

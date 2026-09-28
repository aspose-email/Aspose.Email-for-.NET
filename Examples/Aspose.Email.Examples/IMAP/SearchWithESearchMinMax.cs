// Demonstrates ESEARCH result options (RFC 4731).
//
// Instead of returning every message that matches, the server can report only the
// lowest and/or the highest match - a cheap way to find, for example, the oldest and
// the newest unread message without transferring the full result.

using System;
using Aspose.Email.Clients.Imap;

namespace Aspose.Email.Examples.IMAP
{
    internal static class SearchWithESearchMinMax
    {
        public static void Run()
        {
            using (var client = ClientBuilder.Imap(AuthType.ModernWithDelegatedPermission))
            {
                client.SelectFolder(ImapFolderInfo.InBox);

                if (!client.ESearchSupported)
                {
                    Console.WriteLine("The server does not support ESEARCH.");
                    return;
                }

                var builder = new ImapQueryBuilder();
                builder.HasNoFlags(ImapMessageFlags.IsRead);
                builder.ESearchParameters = ESearchOptions.Min + ESearchOptions.Max;

                var result = client.ListMessages(builder.GetQuery());

                Console.WriteLine($"Lowest and highest unread message ({result.Count} returned):");
                foreach (var info in result)
                    Console.WriteLine($"  #{info.SequenceNumber}  {info.InternalDate:g}  {info.Subject}");
            }
        }
    }
}

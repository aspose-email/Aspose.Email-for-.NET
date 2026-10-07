// Demonstrates searching by the date a message arrived on the server.
//
// Every message has two dates: the Date header, written by the sender's program (and
// sometimes wrong), and the internal date, set by the server when the message was
// stored. InternalDate searches use the latter, which is what "received today" means.
// SentDate searches the Date header instead.

using System;
using Aspose.Email.Clients.Imap;

namespace Aspose.Email.Examples.IMAP
{
    internal static class InternalDateFilter
    {
        public static void Run()
        {
            using (var client = ClientBuilder.Imap(AuthType.ModernWithDelegatedPermission))
            {
                client.SelectFolder(ImapFolderInfo.InBox);

                var builder = new ImapQueryBuilder();
                builder.InternalDate.Since(DateTime.Today.AddDays(-7));

                var messages = client.ListMessages(builder.GetQuery());
                Console.WriteLine($"Received in the last 7 days: {messages.Count} message(s)");

                foreach (var info in messages)
                {
                    Console.WriteLine($"  {info.Subject}");
                    Console.WriteLine($"    received (internal date): {info.InternalDate:g}");
                    Console.WriteLine($"    sent (Date header):       {info.Date:g}");
                }
            }
        }
    }
}

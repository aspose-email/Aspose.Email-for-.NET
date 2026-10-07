// Demonstrates listing a large folder over several connections and compares the time
// with a single connection.
//
// With UseMultiConnection enabled, ListMessages spreads the work over up to
// ConnectionsQuantity parallel connections. Whether that is faster
// depends on the folder size and on how the server handles parallel connections - the
// measured ratio below tells you for your server.

using System;
using System.Diagnostics;
using Aspose.Email.Clients;
using Aspose.Email.Clients.Imap;

namespace Aspose.Email.Examples.IMAP
{
    internal static class ImapListMessagesWithMultiConnection
    {
        public static void Run()
        {
            using (var client = ClientBuilder.Imap(AuthType.ModernWithDelegatedPermission))
            {
                client.SelectFolder(ImapFolderInfo.InBox);
                client.ConnectionsQuantity = 5;

                client.UseMultiConnection = MultiConnectionMode.Enable;
                var watch = Stopwatch.StartNew();
                var multi = client.ListMessages();
                var multiTime = watch.Elapsed;

                client.UseMultiConnection = MultiConnectionMode.Disable;
                watch.Restart();
                var single = client.ListMessages();
                var singleTime = watch.Elapsed;

                Console.WriteLine($"Up to {client.ConnectionsQuantity} connections: {multi.Count} message(s) " +
                                  $"in {multiTime.TotalMilliseconds:F0} ms");
                Console.WriteLine($"One connection:     {single.Count} message(s) " +
                                  $"in {singleTime.TotalMilliseconds:F0} ms");

                if (multiTime.TotalMilliseconds > 0)
                    Console.WriteLine($"Speed-up: {singleTime.TotalMilliseconds / multiTime.TotalMilliseconds:F2}x");
            }
        }
    }
}

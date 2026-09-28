// Demonstrates the two timeouts that keep an IMAP client from hanging forever.
//
// GreetingTimeout bounds how long the client waits for the server's greeting on a new
// connection - keep it short, a server that is up answers almost at once. Timeout
// bounds each operation. When a message cannot be read in time the client throws
// FetchTimeoutException, which you can catch per message and carry on with the rest.

using System;
using Aspose.Email.Clients.Imap;

namespace Aspose.Email.Examples.IMAP
{
    internal static class SetOperationTimeouts
    {
        public static void Run()
        {
            using (var client = ClientBuilder.Imap(AuthType.ModernWithDelegatedPermission))
            {
                client.GreetingTimeout = 5000;  // milliseconds
                client.Timeout = 30000;         // milliseconds, per operation

                Console.WriteLine($"Greeting timeout:  {client.GreetingTimeout} ms");
                Console.WriteLine($"Operation timeout: {client.Timeout} ms\n");

                client.SelectFolder(ImapFolderInfo.InBox);
                var messages = client.ListMessages(ImapFolderInfo.InBox, ImapListFields.IdOnly, 5);

                foreach (var info in messages)
                {
                    try
                    {
                        var message = client.FetchMessage(info.UniqueId);
                        Console.WriteLine($"Fetched   {info.UniqueId}: {message.Subject}");
                    }
                    catch (FetchTimeoutException ex)
                    {
                        Console.WriteLine($"Timed out {info.UniqueId}: {ex.Message}");
                    }
                }
            }
        }
    }
}

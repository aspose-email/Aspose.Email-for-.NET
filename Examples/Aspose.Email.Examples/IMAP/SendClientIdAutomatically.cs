// Demonstrates how to identify your application to the server with the ID command
// (RFC 2971) without calling IntroduceClient yourself.
//
// Set ClientIdentificationInfo and turn on ExchangeIdAutomatically: the client then
// sends ID on its own after connecting, and the server's reply is available from
// ServerIdentificationInfo. Some providers refuse to open folders until the client has
// identified itself; others use the ID only for their logs.

using System;
using Aspose.Email.Clients.Imap;

namespace Aspose.Email.Examples.IMAP
{
    internal static class SendClientIdAutomatically
    {
        public static void Run()
        {
            using (var client = ClientBuilder.Imap(AuthType.ModernWithDelegatedPermission))
            {
                client.ClientIdentificationInfo = new ImapIdentificationInfo
                {
                    Name = "Aspose.Email Examples",
                    Version = "1.0",
                    Vendor = "Example Corp",
                    SupportUrl = "https://www.example.com/support"
                };
                client.ExchangeIdAutomatically = true;

                client.SelectFolder(ImapFolderInfo.InBox);

                if (!client.IdSupported)
                {
                    Console.WriteLine("The server does not support the ID command.");
                    return;
                }

                var server = client.ServerIdentificationInfo;
                if (server == null)
                {
                    Console.WriteLine("The server did not identify itself.");
                    return;
                }

                Console.WriteLine("The server identified itself as:");
                Console.WriteLine($"  name:        {server.Name}");
                Console.WriteLine($"  version:     {server.Version}");
                Console.WriteLine($"  vendor:      {server.Vendor}");
                Console.WriteLine($"  support url: {server.SupportUrl}");
            }
        }
    }
}

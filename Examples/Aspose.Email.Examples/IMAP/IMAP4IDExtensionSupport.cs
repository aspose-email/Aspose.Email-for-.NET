// Demonstrates the ID command (RFC 2971): client and server tell each other who they are.
//
// IntroduceClient sends ID and returns what the server reports about itself - name,
// vendor, version, support address. Without arguments the client sends Aspose.Email's
// default identification; pass an ImapIdentificationInfo to describe your own
// application. SendClientIdAutomatically shows how to have this done on every connect.

using System;
using Aspose.Email.Clients.Imap;

namespace Aspose.Email.Examples.IMAP
{
    internal static class IMAP4IDExtensionSupport
    {
        public static void Run()
        {
            using (var client = ClientBuilder.Imap(AuthType.ModernWithDelegatedPermission))
            {
                client.GetCapabilities();

                if (!client.IdSupported)
                {
                    Console.WriteLine("The server does not support the ID command.");
                    return;
                }

                Print("Reply to the default identification:", client.IntroduceClient());

                var myApplication = new ImapIdentificationInfo
                {
                    Name = "Aspose.Email Examples",
                    Version = "1.0",
                    Vendor = "Example Corp"
                };
                Print("Reply to a custom identification:", client.IntroduceClient(myApplication));
            }
        }

        private static void Print(string title, ImapIdentificationInfo server)
        {
            Console.WriteLine(title);
            Console.WriteLine($"  name:        {server?.Name}");
            Console.WriteLine($"  vendor:      {server?.Vendor}");
            Console.WriteLine($"  version:     {server?.Version}");
            Console.WriteLine($"  support url: {server?.SupportUrl}");
            Console.WriteLine();
        }
    }
}

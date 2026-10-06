// Demonstrates the ENABLE command (RFC 5161), exposed as ClientCapabilities.
//
// Some extensions change what the server sends back, so the server keeps them off until
// the client says it can cope. ClientCapabilities sends ENABLE with the extensions the
// client is prepared to handle and returns the ones the server actually switched on.
// Only list extensions the client really processes, and do it before selecting a folder.

using System;

namespace Aspose.Email.Examples.IMAP
{
    internal static class EnableServerExtensions
    {
        public static void Run()
        {
            using (var client = ClientBuilder.Imap(AuthType.ModernWithDelegatedPermission))
            {
                client.GetCapabilities();

                if (!client.EnableSupported)
                {
                    Console.WriteLine("The server does not support ENABLE.");
                    return;
                }

                var enabled = client.ClientCapabilities("CONDSTORE");

                Console.WriteLine("Requested: CONDSTORE");
                Console.WriteLine("Enabled:   " + (enabled.Length > 0 ? string.Join(", ", enabled) : "(none)"));
            }
        }
    }
}

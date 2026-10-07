// Demonstrates how to list the capabilities the IMAP server announces.
//
// GetCapabilities returns the server's CAPABILITY response: one entry per extension or
// sign-in mechanism, such as IDLE, MOVE, UIDPLUS, CONDSTORE or AUTH=XOAUTH2.
// DetectServerExtensions shows the same information as ready-made *Supported flags.

using System;

namespace Aspose.Email.Examples.IMAP
{
    internal static class RetreivingServerExtensions
    {
        public static void Run()
        {
            using (var client = ClientBuilder.Imap(AuthType.ModernWithDelegatedPermission))
            {
                var capabilities = client.GetCapabilities();

                Console.WriteLine($"{client.Host} announces {capabilities.Length} capabilities:");
                foreach (var capability in capabilities)
                    Console.WriteLine("  " + capability);
            }
        }
    }
}

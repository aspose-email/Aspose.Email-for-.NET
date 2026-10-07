// Demonstrates how to list the extensions the SMTP server announces.
//
// GetCapabilities returns the server's reply to EHLO: one entry per extension, such as
// STARTTLS, AUTH with its mechanisms, SIZE with the largest message accepted, PIPELINING,
// 8BITMIME or DSN. Check it to see, for example, whether delivery notifications or
// pipelining are worth using with this server.

using System;

namespace Aspose.Email.Examples.SMTP
{
    internal static class RetreiveServerExtensions
    {
        public static void Run()
        {
            if (!ClientBuilder.IsSmtpConfigured)
            {
                SmtpExampleInfo.PrintNotConfigured();
                return;
            }

            using (var client = ClientBuilder.Smtp(AuthType.Basic))
            {
                var capabilities = client.GetCapabilities();

                Console.WriteLine($"{client.Host} announces {capabilities.Length} extension(s):");
                foreach (var capability in capabilities)
                    Console.WriteLine("  " + capability);
            }
        }
    }
}

// Demonstrates forwarding a saved message straight from a stream.
//
// The Forward overload that takes a Stream sends the raw message to the given envelope
// recipients without loading it into a MailMessage first - handy for relaying files
// from a folder or an archive. Several recipients go into a MailAddressCollection.

using System;
using System.IO;

namespace Aspose.Email.Examples.SMTP
{
    internal static class ForwardEmailWithoutUsingMailMessage
    {
        public static void Run()
        {
            if (!ClientBuilder.IsSmtpConfigured)
            {
                SmtpExampleInfo.PrintNotConfigured();
                return;
            }

            var emlPath = Data.Email/"test.eml";

            using (var client = ClientBuilder.Smtp(AuthType.Basic))
            using (var stream = File.OpenRead(emlPath))
            {
                var recipients = new MailAddressCollection();
                recipients.Add(client.Username);

                client.Forward(client.Username, recipients, stream);
                Console.WriteLine($"Forwarded {Path.GetFileName(emlPath)} ({stream.Length} bytes) to {recipients}.");
            }
        }
    }
}

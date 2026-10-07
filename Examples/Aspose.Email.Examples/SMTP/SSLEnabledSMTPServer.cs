// Demonstrates how to encrypt the connection to the SMTP server.
//
// There are two ways, tied to the port:
// - SSLExplicit (STARTTLS, usually port 587): the client connects in plain text and
//   upgrades the connection before signing in.
// - SSLImplicit (usually port 465): the connection is encrypted from the first byte.
// SecurityOptions.Auto lets the client work it out; setting the option explicitly
// removes the guesswork. Never send credentials over SecurityOptions.None.

using System;
using Aspose.Email.Clients;

namespace Aspose.Email.Examples.SMTP
{
    internal static class SSLEnabledSMTPServer
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
                client.SecurityOptions = client.Port == 465 ? SecurityOptions.SSLImplicit : SecurityOptions.SSLExplicit;

                client.Send(new MailMessage(client.Username, client.Username, "Sent over TLS", "Body"));
                Console.WriteLine($"Sent to {client.Username} via {client.Host}:{client.Port} " +
                                  $"using {client.SecurityOptions}.");
            }
        }
    }
}

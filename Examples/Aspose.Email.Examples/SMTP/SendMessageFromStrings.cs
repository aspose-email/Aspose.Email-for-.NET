// Demonstrates the shortest way to send a plain-text message.
//
// Send(from, recipients, subject, body) builds the message itself, so there is no
// MailMessage to create. Several recipients go into one comma-separated string. Use a
// MailMessage instead as soon as you need HTML, attachments or extra headers.

using System;

namespace Aspose.Email.Examples.SMTP
{
    internal static class SendMessageFromStrings
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
                client.Send(client.Username, client.Username,
                    "Server status", "All services are running.");

                Console.WriteLine($"Sent a plain-text message to {client.Username}.");
            }
        }
    }
}

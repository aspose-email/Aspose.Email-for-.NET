// Demonstrates the basic way to send a message over SMTP.
//
// Build a MailMessage with a sender, recipients, a subject and a body, then hand it to
// SmtpClient.Send, which connects, signs in and transfers the message; it returns once
// the server has accepted it, and throws an SmtpException if the server refuses.

using System;

namespace Aspose.Email.Examples.SMTP
{
    internal static class SendEmailUsingSMTP
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
                var message = new MailMessage
                {
                    From = client.Username,
                    To = client.Username,
                    Subject = "Hello from Aspose.Email",
                    Body = "This message was sent with SmtpClient.Send."
                };

                client.Send(message);
                Console.WriteLine($"Sent '{message.Subject}' to {message.To} via {client.Host}:{client.Port}.");
            }
        }
    }
}

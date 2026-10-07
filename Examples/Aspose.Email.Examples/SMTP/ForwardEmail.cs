// Demonstrates forwarding a message as it is.
//
// Forward sends an existing message to new recipients without changing it: the original
// From, To and Subject stay in the headers, and only the envelope - who the server
// delivers to - is new. That is how mail is redirected, unlike a "Fwd:" message that
// quotes the original in a new one.

using System;

namespace Aspose.Email.Examples.SMTP
{
    internal static class ForwardEmail
    {
        public static void Run()
        {
            if (!ClientBuilder.IsSmtpConfigured)
            {
                SmtpExampleInfo.PrintNotConfigured();
                return;
            }

            var message = MailMessage.Load(Data.Smtp/"Message.eml");

            using (var client = ClientBuilder.Smtp(AuthType.Basic))
            {
                client.Forward(client.Username, client.Username, message);

                Console.WriteLine($"Forwarded '{message.Subject}' to {client.Username}.");
                Console.WriteLine($"Its headers still name the original recipients: {message.To}");
            }
        }
    }
}

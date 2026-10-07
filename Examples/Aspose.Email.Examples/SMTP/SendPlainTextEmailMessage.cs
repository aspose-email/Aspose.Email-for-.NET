// Demonstrates sending a plain-text message.
//
// Body sets the text; with IsBodyHtml left false the message goes out as text/plain,
// which every mail program displays the same way and spam filters treat kindly.

using System;

namespace Aspose.Email.Examples.SMTP
{
    internal static class SendPlainTextEmailMessage
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
                var message = new MailMessage(client.Username, client.Username)
                {
                    Subject = "Plain-text message",
                    Body = "This is a plain-text body.\r\nLine breaks are kept as they are.",
                    IsBodyHtml = false
                };

                client.Send(message);
                Console.WriteLine($"Sent a plain-text message to {client.Username}.");
            }
        }
    }
}

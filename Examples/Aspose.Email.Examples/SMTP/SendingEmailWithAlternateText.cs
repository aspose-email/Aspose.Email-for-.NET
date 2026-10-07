// Demonstrates an HTML message with a plain-text alternative.
//
// Alternate views put several versions of the same content into one message
// (multipart/alternative); each mail program shows the richest one it can display.
// Here the body is HTML and the alternate view adds plain text for programs, screen
// readers and filters that prefer it. SendEmailWithAlternateText shows the opposite
// arrangement.

using System;

namespace Aspose.Email.Examples.SMTP
{
    internal static class SendingEmailWithAlternateText
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
                    Subject = "HTML with a plain-text alternative",
                    HtmlBody = "<p>Your order <b>A-1001</b> has shipped.</p>"
                };

                message.AlternateViews.Add(AlternateView.CreateAlternateViewFromString(
                    "Your order A-1001 has shipped."));

                client.Send(message);
                Console.WriteLine($"Sent to {client.Username} with {message.AlternateViews.Count} alternate view(s).");
            }
        }
    }
}

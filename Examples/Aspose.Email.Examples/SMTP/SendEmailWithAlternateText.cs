// Demonstrates a plain-text message with an HTML alternative.
//
// The body is plain text and an alternate view of type text/html carries the formatted
// version; mail programs that can show HTML pick it, the rest show the text. This suits
// messages generated from plain text first. SendingEmailWithAlternateText shows the
// opposite arrangement.

using System;
using System.Text;

namespace Aspose.Email.Examples.SMTP
{
    internal static class SendEmailWithAlternateText
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
                    Subject = "Plain text with an HTML alternative",
                    Body = "Build 1234 passed. Details: https://ci.example.com/1234"
                };

                message.AlternateViews.Add(AlternateView.CreateAlternateViewFromString(
                    "<p>Build <b>1234</b> <span style=\"color:green\">passed</span>. " +
                    "<a href=\"https://ci.example.com/1234\">Details</a></p>",
                    Encoding.UTF8, "text/html"));

                client.Send(message);
                Console.WriteLine($"Sent to {client.Username} with {message.AlternateViews.Count} alternate view(s).");
            }
        }
    }
}

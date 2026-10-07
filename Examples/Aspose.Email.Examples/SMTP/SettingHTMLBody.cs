// Demonstrates sending an HTML message.
//
// Setting HtmlBody makes the message text/html, so the recipient sees formatting, links
// and colours. Keep the markup simple and the styles inline - mail programs ignore most
// of what a browser supports. SendingEmailWithAlternateText adds a plain-text version
// for programs that do not show HTML.

using System;

namespace Aspose.Email.Examples.SMTP
{
    internal static class SettingHTMLBody
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
                    Subject = "HTML message",
                    HtmlBody = "<html><body>" +
                               "<h2 style=\"color:#2b579a\">Monthly report</h2>" +
                               "<p>Sales are <b>up 12%</b>. <a href=\"https://www.example.com/report\">Read more</a>.</p>" +
                               "</body></html>"
                };

                client.Send(message);
                Console.WriteLine($"Sent an HTML message to {client.Username} (IsBodyHtml = {message.IsBodyHtml}).");
            }
        }
    }
}

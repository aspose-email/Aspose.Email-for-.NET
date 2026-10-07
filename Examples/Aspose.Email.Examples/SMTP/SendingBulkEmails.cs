// Demonstrates sending several different messages in one call.
//
// Send accepts a MailMessageCollection (or any IEnumerable<MailMessage>) and sends the
// messages one after another over the same connection. If some of them fail, the
// SmtpException it throws lists what was sent and what was not - see
// HandleSmtpSendErrors.

using System;

namespace Aspose.Email.Examples.SMTP
{
    internal static class SendingBulkEmails
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
                var self = client.Username;

                var messages = new MailMessageCollection
                {
                    new MailMessage(self, self, "Invoice 1001", "Your invoice is attached."),
                    new MailMessage(self, self, "Invoice 1002", "Your invoice is attached."),
                    new MailMessage(self, self, "Invoice 1003", "Your invoice is attached.")
                };

                client.Send(messages);
                Console.WriteLine($"Sent {messages.Count} message(s) to {self} in one call.");
            }
        }
    }
}

// Demonstrates sending one message to several To, Cc and Bcc recipients.
//
// To and Cc recipients appear in the message headers; Bcc recipients receive the message
// but are left out of the headers, so the others do not see them. The server gets every
// address as a separate envelope recipient.
//
// To stay within your own mailbox the example uses sub-addresses of your address
// (user+tag@domain), which Gmail, Outlook.com, Microsoft 365 and many other providers
// deliver to the same mailbox. If yours does not, put real addresses here.

using System;

namespace Aspose.Email.Examples.SMTP
{
    internal static class MultipleRecipients
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
                var self = new MailAddress(client.Username);

                var message = new MailMessage
                {
                    From = self,
                    Subject = "Message to several recipients",
                    Body = "Check the To and Cc lines; the Bcc recipient is not listed."
                };

                message.To.Add(SubAddress(self, "to1"));
                message.To.Add(SubAddress(self, "to2"));
                message.CC.Add(SubAddress(self, "cc1"));
                message.Bcc.Add(SubAddress(self, "bcc1"));

                client.Send(message);

                Console.WriteLine("Sent to:");
                Console.WriteLine($"  To:  {message.To}");
                Console.WriteLine($"  Cc:  {message.CC}");
                Console.WriteLine($"  Bcc: {message.Bcc}");
            }
        }

        private static string SubAddress(MailAddress address, string tag)
        {
            return $"{address.User}+{tag}@{address.Host}";
        }
    }
}

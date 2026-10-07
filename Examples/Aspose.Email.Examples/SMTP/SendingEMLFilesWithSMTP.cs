// Demonstrates sending a message that was saved as an .eml file earlier.
//
// MailMessage.Load restores the message completely - body, attachments and headers - so
// it can be sent again as it is or after changing its recipients. ExportAsEML shows how
// such a file is created; ForwardEmailWithoutUsingMailMessage sends a file without
// loading it.

using System;

namespace Aspose.Email.Examples.SMTP
{
    internal static class SendingEMLFilesWithSMTP
    {
        public static void Run()
        {
            if (!ClientBuilder.IsSmtpConfigured)
            {
                SmtpExampleInfo.PrintNotConfigured();
                return;
            }

            var emlPath = Data.Smtp/"Message.eml";
            var message = MailMessage.Load(emlPath);
            Console.WriteLine($"Loaded '{message.Subject}', originally to {message.To}");

            using (var client = ClientBuilder.Smtp(AuthType.Basic))
            {
                // The file is addressed to sample recipients; send it to yourself instead.
                message.From = client.Username;
                message.To.Clear();
                message.CC.Clear();
                message.To.Add(client.Username);

                client.Send(message);
                Console.WriteLine($"Sent to {client.Username}.");
            }
        }
    }
}

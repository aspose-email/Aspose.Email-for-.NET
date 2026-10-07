// Demonstrates loading an .eml file with load options and sending it.
//
// EmlLoadOptions controls how the file is parsed - PreserveTnefAttachments, for example,
// decides whether the files inside a winmail.dat (TNEF) attachment are extracted or the
// attachment is kept as it is. After loading, the message is an ordinary MailMessage
// whose headers can be changed before sending.

using System;

namespace Aspose.Email.Examples.SMTP
{
    internal static class ImportEML
    {
        public static void Run()
        {
            if (!ClientBuilder.IsSmtpConfigured)
            {
                SmtpExampleInfo.PrintNotConfigured();
                return;
            }

            var emlPath = Data.Smtp/"test.eml";
            var message = MailMessage.Load(emlPath, new EmlLoadOptions { PreserveTnefAttachments = true });

            Console.WriteLine($"Loaded {emlPath}");
            Console.WriteLine($"  subject:     {message.Subject}");
            Console.WriteLine($"  from:        {message.From}");
            Console.WriteLine($"  attachments: {message.Attachments.Count}");

            using (var client = ClientBuilder.Smtp(AuthType.Basic))
            {
                // The file is addressed to sample recipients; send it to yourself instead.
                message.From = client.Username;
                message.To.Clear();
                message.CC.Clear();
                message.Bcc.Clear();
                message.To.Add(client.Username);

                client.Send(message);
                Console.WriteLine($"\nSent to {client.Username}.");
            }
        }
    }
}

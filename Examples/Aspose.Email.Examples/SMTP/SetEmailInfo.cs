// Demonstrates the message properties that tell the recipient how to treat a message:
// its date, priority and sensitivity.
//
// Priority is written to the priority headers that mail programs show as a flag or an
// exclamation mark; Sensitivity (Personal, Private, Company-Confidential) is shown as a notice in
// Outlook. Neither changes how the server delivers the message.

using System;

namespace Aspose.Email.Examples.SMTP
{
    internal static class SetEmailInfo
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
                var message = new MailMessage(client.Username, client.Username, "Quarterly results", "Please review.")
                {
                    Date = DateTime.Now,
                    Priority = MailPriority.High,
                    Sensitivity = MailSensitivity.CompanyConfidential
                };

                client.Send(message);

                Console.WriteLine($"Sent to {client.Username}:");
                Console.WriteLine($"  date:        {message.Date}");
                Console.WriteLine($"  priority:    {message.Priority}");
                Console.WriteLine($"  sensitivity: {message.Sensitivity}");
            }
        }
    }
}

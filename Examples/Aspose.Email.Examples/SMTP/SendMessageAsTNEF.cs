// Demonstrates sending a message in Outlook's TNEF format (winmail.dat).
//
// With UseTnef set, the client packs the message into a TNEF attachment, which keeps
// Outlook-specific properties - rich text, voting buttons, custom forms - intact between
// Outlook and Exchange. Programs other than Outlook usually just show a winmail.dat
// attachment, so use it only when the recipients run Outlook.
// PreserveTnefAttachments keeps an incoming TNEF attachment as it is while loading.

using System;

namespace Aspose.Email.Examples.SMTP
{
    internal static class SendMessageAsTNEF
    {
        public static void Run()
        {
            if (!ClientBuilder.IsSmtpConfigured)
            {
                SmtpExampleInfo.PrintNotConfigured();
                return;
            }

            var message = MailMessage.Load(Data.Email/"Message.eml", new EmlLoadOptions { PreserveTnefAttachments = true });

            using (var client = ClientBuilder.Smtp(AuthType.Basic))
            {
                message.From = client.Username;
                message.To.Clear();
                message.CC.Clear();
                message.Bcc.Clear();
                message.To.Add(client.Username);
                message.Subject = "Sent as TNEF";
                message.Date = DateTime.Now;

                client.UseTnef = true;
                client.Send(message);

                Console.WriteLine($"Sent to {client.Username} in TNEF format.");
            }
        }
    }
}

// Demonstrates how to build an email that asks for a delivery notification and a
// read receipt, and how to send it via SmtpClient.

using System;

namespace Aspose.Email.Examples.Email
{
    internal static class RequestReadReceipt
    {
        public static void Run()
        {
            var eml = new MailMessage
            {
                From = "sender@sender.com",
                To = "receiver@receiver.com",
                Subject = "Message requesting a read receipt",
                HtmlBody = "<html><body>This is the Html body</body></html>",

                // Asks the receiving server to confirm successful delivery.
                DeliveryNotificationOptions = DeliveryNotificationOptions.OnSuccess
            };

            // Asks the recipient's mail client to confirm that the message was opened.
            eml.Headers.Add("Return-Receipt-To", "sender@sender.com");
            eml.Headers.Add("Disposition-Notification-To", "sender@sender.com");

            var outputPath = Data.Out/"RequestReadReceipt_out.eml";
            eml.Save(outputPath, SaveOptions.DefaultEml);
            Console.WriteLine($"Message saved to {outputPath}");

            // Sending needs a real server, so it only runs once one is configured in
            // clientsettings.json. The message above is complete either way.
            if (!ClientBuilder.IsSmtpConfigured)
            {
                Console.WriteLine("Set Smtp.HostName in clientsettings.json to also send this message.");
                return;
            }

            using (var client = ClientBuilder.Smtp(AuthType.Basic))
            {
                client.Send(eml);
                Console.WriteLine("Message sent.");
            }
        }
    }
}

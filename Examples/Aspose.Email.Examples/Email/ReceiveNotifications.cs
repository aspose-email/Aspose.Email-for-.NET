// Demonstrates how to request delivery notifications for both success and failure,
// together with a read receipt.

using System;

namespace Aspose.Email.Examples.Email
{
    internal static class ReceiveNotifications
    {
        public static void Run()
        {
            var eml = new MailMessage
            {
                From = "sender@sender.com",
                To = "receiver@receiver.com",
                Subject = "the subject of the message",

                // Ask the server to report both a successful delivery and a failure.
                DeliveryNotificationOptions =
                    DeliveryNotificationOptions.OnSuccess | DeliveryNotificationOptions.OnFailure
            };

            // Ask the recipient's mail client for a read receipt.
            eml.Headers.Add("Disposition-Notification-To", "sender@sender.com");

            var outputPath = Data.Out/"ReceiveNotifications_out.eml";
            eml.Save(outputPath, SaveOptions.DefaultEml);

            Console.WriteLine($"Notification options: {eml.DeliveryNotificationOptions}");
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

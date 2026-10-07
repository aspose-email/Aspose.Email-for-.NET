// Demonstrates how to request delivery status notifications (DSN, RFC 3461).
//
// DeliveryNotificationOptions asks the mail servers along the way to report back to the
// sender: on successful delivery, on failure, and/or when delivery is delayed. A server
// that does not support the DSN extension ignores the request, and many providers only
// honour failure reports - so treat a missing notification as "unknown", not "failed".
// A read receipt, which the recipient's program sends, is a different mechanism.

using System;

namespace Aspose.Email.Examples.SMTP
{
    internal static class DeliveryNotifications
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
                var message = new MailMessage(client.Username, client.Username, "Delivery report requested", "Body")
                {
                    DeliveryNotificationOptions =
                        DeliveryNotificationOptions.OnSuccess |
                        DeliveryNotificationOptions.OnFailure |
                        DeliveryNotificationOptions.Delay
                };

                client.Send(message);
                Console.WriteLine($"Sent to {client.Username}, requesting: {message.DeliveryNotificationOptions}.");
                Console.WriteLine("Delivery reports, if the servers support DSN, arrive at the sender address.");
            }
        }
    }
}

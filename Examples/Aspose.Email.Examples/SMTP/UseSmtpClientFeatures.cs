// Demonstrates the basic SmtpClient settings that decide how and how long it tries.
//
// DeliveryMethod chooses between sending over the network (the default) and writing to
// a pickup directory - see SaveToPickupDirectory. GreetingTimeout bounds the wait for
// the server's greeting on a new connection; Timeout bounds each operation. Both are in
// milliseconds.

using System;
using Aspose.Email.Clients.Smtp;

namespace Aspose.Email.Examples.SMTP
{
    internal static class UseSmtpClientFeatures
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
                client.DeliveryMethod = SmtpDeliveryMethod.Network;
                client.GreetingTimeout = 5000;
                client.Timeout = 30000;

                Console.WriteLine($"Delivery method:   {client.DeliveryMethod}");
                Console.WriteLine($"Greeting timeout:  {client.GreetingTimeout} ms");
                Console.WriteLine($"Operation timeout: {client.Timeout} ms");

                client.Send(new MailMessage(client.Username, client.Username, "Client settings test", "Body"));
                Console.WriteLine($"\nSent to {client.Username}.");
            }
        }
    }
}

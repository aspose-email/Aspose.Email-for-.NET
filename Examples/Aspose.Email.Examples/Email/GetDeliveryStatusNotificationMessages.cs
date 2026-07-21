// Demonstrates how to load an email message and inspect its bounce/delivery status
// notification details using CheckBounced.

using System;

namespace Aspose.Email.Examples.Email
{
    internal static class GetDeliveryStatusNotificationMessages
    {
        public static void Run()
        {
            var eml = MailMessage.Load(Data.Email/"failed1.msg");
            var result = eml.CheckBounced();

            Console.WriteLine($"IsBounced: {result.IsBounced}");
            Console.WriteLine($"Action: {result.Action}");
            Console.WriteLine($"Recipient: {result.Recipient}");
            Console.WriteLine($"Reason: {result.Reason}");
            Console.WriteLine($"Status: {result.Status}");
            Console.WriteLine($"OriginalMessage ToAddress 1: {result.OriginalMessage.To[0].Address}");
        }
    }
}

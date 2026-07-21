// Demonstrates how to read the delivery and read receipt timestamps that Outlook
// records for each recipient of a message.

using System;
using Aspose.Email.Mapi;

namespace Aspose.Email.Examples.MAPI
{
    internal static class RetrieveReadAndDeliveryReceiptInformation
    {
        public static void Run()
        {
            var msg = MapiMessage.Load(Data.Mapi/"TestMessage.msg");

            Console.WriteLine($"Message: {msg.Subject}");
            Console.WriteLine($"Recipients: {msg.Recipients.Count}\n");

            foreach (var recipient in msg.Recipients)
            {
                Console.WriteLine($"Recipient:     {recipient.DisplayName}");
                Console.WriteLine($"Delivery time: {GetTime(recipient, MapiPropertyTag.PR_RECIPIENT_TRACKSTATUS_TIME_DELIVERY)}");
                Console.WriteLine($"Read time:     {GetTime(recipient, MapiPropertyTag.PR_RECIPIENT_TRACKSTATUS_TIME_READ)}");
                Console.WriteLine();
            }
        }

        // Receipt times are tracked per recipient, and are simply absent when the message
        // was never delivered or never read - the indexer returns null in that case.
        private static string GetTime(MapiRecipient recipient, long tag)
        {
            var property = recipient.Properties[tag];
            return property == null ? "(not set)" : property.GetDateTime().ToString("u");
        }
    }
}

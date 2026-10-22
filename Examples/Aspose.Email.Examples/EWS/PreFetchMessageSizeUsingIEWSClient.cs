// Demonstrates reading message sizes before downloading anything.
//
// Every ExchangeMessageInfo from a listing carries the message size, so a client can
// decide up front what to download - skip huge messages on a metered connection, or
// fetch them later in the background.

using System;
using System.Linq;

namespace Aspose.Email.Examples.EWS
{
    internal static class PreFetchMessageSizeUsingIEWSClient
    {
        public static void Run()
        {
            const long largeMessage = 1024 * 1024;

            using (var client = ClientBuilder.Ews(AuthType.ModernWithAppPermission))
            {
                var messages = client.ListMessages(client.MailboxInfo.InboxUri, 50);

                Console.WriteLine($"{messages.Count} message(s), {messages.Sum(info => info.Size) / 1024} KB in total.");
                Console.WriteLine("Larger than 1 MB:");

                foreach (var info in messages.Where(info => info.Size > largeMessage))
                    Console.WriteLine($"  {info.Size / 1024,8} KB  {info.Subject}");
            }
        }
    }
}

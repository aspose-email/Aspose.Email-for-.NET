// Demonstrates listing the messages of a Graph folder and fetching one in full.
//
// ListMessages returns MessageInfo - the summary Graph sends back for a listing.
// FetchMessage turns one of those into a MapiMessage with body and attachments.

using System;
using Aspose.Email.Clients.Graph;

namespace Aspose.Email.Examples.Graph
{
    internal static class ListMessagesFromGraph
    {
        public static void Run()
        {
            if (!ClientBuilder.IsGraphConfigured)
            {
                GraphExampleInfo.PrintNotConfigured();
                return;
            }

            var outputDir = Data.OutSub("GraphMessages");

            using (var client = ClientBuilder.Graph(AuthType.ModernWithAppPermission))
            {
                var messages = client.ListMessages(KnownFolders.Inbox, null);
                Console.WriteLine($"{messages.Count} message(s) in the Inbox");

                var index = 0;

                foreach (MessageInfo messageInfo in messages)
                {
                    Console.WriteLine($"  {messageInfo.Subject}");
                    Console.WriteLine($"    from:    {messageInfo.From}");
                    Console.WriteLine($"    sent:    {messageInfo.Date}");
                    Console.WriteLine($"    size:    {messageInfo.Size:N0} bytes");
                    Console.WriteLine($"    item id: {messageInfo.ItemId}");

                    // Only the first few are downloaded in full - a listing is cheap,
                    // a fetch is a request per message.
                    if (index < 3)
                    {
                        var message = client.FetchMessage(messageInfo.ItemId);
                        message.Save(outputDir/$"message-{index}.msg");

                        Console.WriteLine($"    body:    {message.Body.Length} character(s), " +
                                          $"{message.Attachments.Count} attachment(s)");
                    }

                    index++;
                }

                Console.WriteLine($"\nSaved the first message(s) to {outputDir}");
            }
        }
    }
}

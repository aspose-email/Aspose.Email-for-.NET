// Demonstrates listing the conversations in a folder.
//
// FindConversations groups the messages of a folder the way Outlook's conversation view
// does and returns one ExchangeConversation per thread: its topic, how many messages it
// has and how many are unread, who took part, and its flag status. The ConversationId
// is what the conversation-wide operations (copy, move, delete, mark read) take.

using System;
using System.Linq;

namespace Aspose.Email.Examples.EWS
{
    internal static class FindConversationsOnExchangeServer
    {
        public static void Run()
        {
            using (var client = ClientBuilder.Ews(AuthType.ModernWithAppPermission))
            {
                var conversations = client.FindConversations(client.MailboxInfo.InboxUri);
                Console.WriteLine($"{conversations.Length} conversation(s) in the Inbox, the 10 latest:");

                foreach (var conversation in conversations.OrderByDescending(c => c.LastDeliveryTime).Take(10))
                {
                    Console.WriteLine($"  {conversation.ConversationTopic}");
                    Console.WriteLine($"    {conversation.MessageCount} message(s), " +
                                      $"{conversation.UnreadCount} unread, " +
                                      $"flag: {conversation.FlagStatus}, last: {conversation.LastDeliveryTime:g}");
                }
            }
        }
    }
}

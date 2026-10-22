// Demonstrates fetching extra MAPI properties as part of a message listing.
//
// The ListMessages overload that takes property descriptors returns those properties
// with each summary, in ExchangeMessageInfo.Properties - so values such as the
// conversation topic or importance are available without downloading the messages.
// Any property from KnownPropertyList, or a custom named property, can be requested.

using System;
using Aspose.Email.Clients.Exchange;
using Aspose.Email.Mapi;

namespace Aspose.Email.Examples.EWS
{
    internal static class RetrieveExtraParametersAsSummaryInformation
    {
        public static void Run()
        {
            var extraProperties = new PropertyDescriptor[]
            {
                KnownPropertyList.ConversationTopic,
                KnownPropertyList.Importance,
                KnownPropertyList.LastModificationTime
            };

            using (var client = ClientBuilder.Ews(AuthType.ModernWithAppPermission))
            {
                var messages = client.ListMessages(client.MailboxInfo.InboxUri, ExchangeListMessagesOptions.None,
                    extraProperties);

                foreach (var info in messages)
                {
                    Console.WriteLine(info.Subject);
                    var topic = info.Properties[KnownPropertyList.ConversationTopic]?.GetString();
                    Console.WriteLine($"  topic:         {topic}");
                    Console.WriteLine($"  importance:    {info.Properties[KnownPropertyList.Importance]?.GetInt32()}");
                    var modified = info.Properties[KnownPropertyList.LastModificationTime]?.GetDateTime();
                    Console.WriteLine($"  last modified: {modified:g}");
                }
            }
        }
    }
}

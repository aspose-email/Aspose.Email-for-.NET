// Demonstrates fetching several items together with extended properties.
//
// FetchItems takes an EwsFetchItems parameter object: the item URIs, the extended
// properties to include, and whether attachments are downloaded. One request returns
// all the items as MapiMessage objects. The example tags three messages with a custom
// property in a uniquely named folder, fetches them back and deletes the folder.

using System;
using System.Linq;
using Aspose.Email.Clients.Exchange.WebService.Models;
using Aspose.Email.Mapi;

namespace Aspose.Email.Examples.EWS
{
    internal static class FetchItemWithExtendedProperties
    {
        public static void Run()
        {
            var folderName = "Aspose-" + Guid.NewGuid().ToString("N").Substring(0, 8);
            var property = new PidNamePropertyDescriptor("TicketId", PropertyDataType.String,
                KnownPropertySets.PublicStrings);

            using (var client = ClientBuilder.Ews(AuthType.ModernWithAppPermission))
            {
                var folder = client.CreateFolder(client.MailboxInfo.InboxUri, folderName);

                try
                {
                    var uris = Enumerable.Range(1, 3).Select(i =>
                    {
                        var message = new MapiMessage("sender@example.com", "receiver@example.com",
                            $"Ticket {i}", "Body");
                        message.SetProperty(property, $"T-{100 + i}");
                        return client.AppendMessage(folder.Uri, message, true);
                    }).ToList();

                    var items = client.FetchItems(EwsFetchItems.Create()
                        .AddUris(uris)
                        .AddExtendedProperty(property)
                        .SkipAttachments());

                    foreach (var item in items)
                        Console.WriteLine($"{item.Subject}: TicketId = {item.NamedProperties[property]?.GetString()}");
                }
                finally
                {
                    client.DeleteFolder(folder.Uri, true);
                }
            }
        }
    }
}

// Demonstrates storing your own data on a message as an extended (named) MAPI property.
//
// A PidNamePropertyDescriptor defines a custom property by name, type and property set.
// Set it on a MapiMessage, upload the message, and request the same descriptor when
// fetching to get the value back - Exchange only returns custom properties you ask for.
// The example works in a uniquely named folder and deletes it at the end.

using System;
using Aspose.Email.Mapi;

namespace Aspose.Email.Examples.EWS
{
    internal static class WorkingWithExtendedAttributesOfMessages
    {
        public static void Run()
        {
            var folderName = "Aspose-" + Guid.NewGuid().ToString("N").Substring(0, 8);
            var property = new PidNamePropertyDescriptor("OrderNumber", PropertyDataType.String,
                KnownPropertySets.PublicStrings);

            using (var client = ClientBuilder.Ews(AuthType.ModernWithAppPermission))
            {
                var folder = client.CreateFolder(client.MailboxInfo.InboxUri, folderName);

                try
                {
                    var message = new MapiMessage("sender@example.com", "receiver@example.com",
                        "Order confirmation", "Body");
                    message.SetProperty(property, "A-1001");

                    var uri = client.AppendMessage(folder.Uri, message, true);

                    var fetched = client.FetchItem(uri, new PropertyDescriptor[] { property });
                    var orderNumber = fetched.NamedProperties[property].GetString();
                    Console.WriteLine($"'{fetched.Subject}' carries OrderNumber = {orderNumber}");
                }
                finally
                {
                    client.DeleteFolder(folder.Uri, true);
                }
            }
        }
    }
}

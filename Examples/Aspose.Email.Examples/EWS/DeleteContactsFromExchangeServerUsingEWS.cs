// Demonstrates deleting a contact.
//
// Contacts are items like messages, so DeleteItem removes them by id - Id.EWSId of the
// Contact - with the usual DeletionOptions. The example creates two contacts in a
// uniquely named contacts folder, deletes one by name and deletes the folder at the end.

using System;
using Aspose.Email.Clients.Exchange;
using Aspose.Email.Clients.Exchange.WebService;
using Aspose.Email.PersonalInfo;

namespace Aspose.Email.Examples.EWS
{
    internal static class DeleteContactsFromExchangeServerUsingEWS
    {
        public static void Run()
        {
            const string contactToDelete = "John Teddy";
            var folderName = "Aspose-" + Guid.NewGuid().ToString("N").Substring(0, 8);

            using (var client = ClientBuilder.Ews(AuthType.ModernWithAppPermission))
            {
                var folder = client.CreateFolder(client.MailboxInfo.ContactsUri, folderName, ExchangeFolderType.Contact);

                try
                {
                    client.CreateContact(folder.Uri, new Contact { DisplayName = contactToDelete });
                    client.CreateContact(folder.Uri, new Contact { DisplayName = "Mary Keep" });

                    foreach (var contact in client.GetContacts(folder.Uri))
                    {
                        if (contact.DisplayName == contactToDelete)
                        {
                            client.DeleteItem(contact.Id.EWSId, DeletionOptions.DeletePermanently);
                            Console.WriteLine($"Deleted {contact.DisplayName}.");
                        }
                    }

                    Console.WriteLine("Left in the folder:");
                    foreach (var contact in client.GetContacts(folder.Uri))
                        Console.WriteLine($"  {contact.DisplayName}");
                }
                finally
                {
                    client.DeleteFolder(folder.Uri, true);
                }
            }
        }
    }
}

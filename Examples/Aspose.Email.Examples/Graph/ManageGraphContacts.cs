// Demonstrates the contact side of Graph: the contact folders, and creating, reading
// and updating contacts inside them.

using System;
using System.Linq;
using Aspose.Email.Clients.Graph;
using Aspose.Email.Mapi;

namespace Aspose.Email.Examples.Graph
{
    internal static class ManageGraphContacts
    {
        public static void Run()
        {
            if (!ClientBuilder.IsGraphConfigured)
            {
                GraphExampleInfo.PrintNotConfigured();
                return;
            }

            using (var client = ClientBuilder.Graph(AuthType.ModernWithAppPermission))
            {
                // The default contacts folder of the mailbox.
                var defaultFolder = client.GetContactFolder();
                Console.WriteLine($"Default contact folder: {defaultFolder.DisplayName} ({defaultFolder.ItemId})");

                var folders = client.ListContactFolders(null);
                Console.WriteLine($"{folders.Count} contact folder(s):");
                foreach (var folder in folders)
                    Console.WriteLine($"  {folder.DisplayName}");

                var contact = new MapiContact("Jane Smith", "jane.smith@example.com", "XYZ Corporation", "777-888-999");

                var created = client.CreateContact(defaultFolder.ItemId, contact);
                Console.WriteLine($"\nCreated: {created.NameInfo.DisplayName}");

                var contacts = client.ListContacts(defaultFolder.ItemId, null);
                Console.WriteLine($"{contacts.Count} contact(s) in the folder:");

                foreach (var item in contacts.Take(10))
                {
                    Console.WriteLine($"  {item.NameInfo.DisplayName} " +
                                      $"<{item.ElectronicAddresses.Email1.EmailAddress}>");
                }

                // Reading one back, changing it, and pushing the change.
                var fetched = client.FetchContact(created.ItemId);
                Console.WriteLine($"\nFetched: {fetched.NameInfo.DisplayName}");

                fetched.Telephones.PrimaryTelephoneNumber = "888-888-999";
                var updated = client.UpdateContact(fetched);
                Console.WriteLine($"Updated phone to: {updated.Telephones.PrimaryTelephoneNumber}");

                client.Delete(created.ItemId);
                Console.WriteLine("Deleted the contact.");
            }
        }
    }
}

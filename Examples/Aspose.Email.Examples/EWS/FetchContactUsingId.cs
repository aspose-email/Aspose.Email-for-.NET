// Demonstrates fetching a single contact by its id.
//
// Every Contact carries its Exchange id in Id.EWSId. Keep that id - in your own database,
// for example - and GetContact fetches the current version of the contact later without
// listing the folder again.

using System;
using System.Linq;

namespace Aspose.Email.Examples.EWS
{
    internal static class FetchContactUsingId
    {
        public static void Run()
        {
            using (var client = ClientBuilder.Ews(AuthType.ModernWithAppPermission))
            {
                var first = client.GetContacts(client.MailboxInfo.ContactsUri).FirstOrDefault();
                if (first == null)
                {
                    Console.WriteLine("The Contacts folder is empty.");
                    return;
                }

                var id = first.Id.EWSId;
                Console.WriteLine($"Stored id: {id}");

                var contact = client.GetContact(id);
                Console.WriteLine($"Fetched:   {contact.DisplayName}, {contact.CompanyName}, " +
                                  $"{contact.EmailAddresses.FirstOrDefault()?.Address}");
            }
        }
    }
}

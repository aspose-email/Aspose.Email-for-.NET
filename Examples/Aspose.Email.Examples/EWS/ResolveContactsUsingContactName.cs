// Demonstrates resolving a partial name to contacts, like "Check Names" in Outlook.
//
// ResolveContacts searches the user's contacts and the directory for entries matching
// an ambiguous name or address fragment and returns them as Contact objects;
// ExchangeListContactsOptions.FetchPhoto also loads their photos.

using System;
using System.Linq;
using Aspose.Email.Clients.Exchange.WebService;

namespace Aspose.Email.Examples.EWS
{
    internal static class ResolveContactsUsingContactName
    {
        public static void Run()
        {
            const string name = "john";

            using (var client = ClientBuilder.Ews(AuthType.ModernWithAppPermission))
            {
                var contacts = client.ResolveContacts(name, ExchangeListContactsOptions.FetchPhoto);

                Console.WriteLine($"'{name}' resolves to {contacts.Length} contact(s):");
                foreach (var contact in contacts)
                {
                    Console.WriteLine($"  {contact.DisplayName} <{contact.EmailAddresses.FirstOrDefault()?.Address}>" +
                                      (contact.Photo != null ? "  [photo]" : ""));
                }
            }
        }
    }
}

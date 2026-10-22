// Demonstrates listing the contacts of a mailbox.
//
// GetContacts returns the contacts of a folder as Contact objects - names, e-mail
// addresses, phone numbers, company and so on. ListContactsAsMapiContacts shows the
// same data as Outlook MapiContact objects.

using System;
using System.Linq;

namespace Aspose.Email.Examples.EWS
{
    internal static class GettingContactsUsingEWS
    {
        public static void Run()
        {
            using (var client = ClientBuilder.Ews(AuthType.ModernWithAppPermission))
            {
                var contacts = client.GetContacts(client.MailboxInfo.ContactsUri);
                Console.WriteLine($"{contacts.Length} contact(s):");

                foreach (var contact in contacts)
                {
                    var email = contact.EmailAddresses.FirstOrDefault()?.Address;
                    Console.WriteLine($"  {contact.DisplayName} <{email}>");
                }
            }
        }
    }
}

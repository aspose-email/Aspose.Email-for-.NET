// Demonstrates listing contacts as MapiContact objects.
//
// GetContacts returns the vCard-style Contact model; ListContacts returns MapiContact,
// the Outlook model with its property sets - names, e-mail addresses, company, phones -
// which can be saved straight to .msg or .vcf. Extra MAPI properties can be requested
// through the overload that takes property descriptors.

using System;

namespace Aspose.Email.Examples.EWS
{
    internal static class ListContactsAsMapiContacts
    {
        public static void Run()
        {
            using (var client = ClientBuilder.Ews(AuthType.ModernWithAppPermission))
            {
                var contacts = client.ListContacts(client.MailboxInfo.ContactsUri);
                Console.WriteLine($"{contacts.Length} contact(s) in the Contacts folder:");

                foreach (var contact in contacts)
                {
                    Console.WriteLine($"  {contact.NameInfo?.DisplayName}");
                    Console.WriteLine($"    e-mail:  {contact.ElectronicAddresses?.Email1?.EmailAddress}");
                    Console.WriteLine($"    company: {contact.ProfessionalInfo?.CompanyName}");
                }
            }
        }
    }
}

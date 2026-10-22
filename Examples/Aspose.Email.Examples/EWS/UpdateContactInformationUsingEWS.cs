// Demonstrates changing an existing contact.
//
// Fetch the contact, change its properties and pass it to UpdateContact; Exchange
// writes back the changes. The example creates a contact in a uniquely named contacts
// folder, updates it and deletes the folder at the end.

using System;
using System.Linq;
using Aspose.Email.Clients.Exchange;
using Aspose.Email.PersonalInfo;

namespace Aspose.Email.Examples.EWS
{
    internal static class UpdateContactInformationUsingEWS
    {
        public static void Run()
        {
            var folderName = "Aspose-" + Guid.NewGuid().ToString("N").Substring(0, 8);

            using (var client = ClientBuilder.Ews(AuthType.ModernWithAppPermission))
            {
                var folder = client.CreateFolder(client.MailboxInfo.ContactsUri, folderName, ExchangeFolderType.Contact);

                try
                {
                    client.CreateContact(folder.Uri, new Contact { DisplayName = "David Chen", JobTitle = "Analyst" });

                    var contact = client.GetContacts(folder.Uri).First();
                    Console.WriteLine($"Before: {contact.DisplayName}, {contact.JobTitle}");

                    contact.JobTitle = "Senior Analyst";
                    contact.CompanyName = "Example Corp";
                    client.UpdateContact(contact);

                    var updated = client.GetContact(contact.Id.EWSId);
                    Console.WriteLine($"After:  {updated.DisplayName}, {updated.JobTitle} at {updated.CompanyName}");
                }
                finally
                {
                    client.DeleteFolder(folder.Uri, true);
                }
            }
        }
    }
}

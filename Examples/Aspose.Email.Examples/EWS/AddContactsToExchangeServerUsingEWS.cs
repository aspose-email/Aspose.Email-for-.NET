// Demonstrates creating a contact with phone numbers, related people, web sites and an
// e-mail address.
//
// Contact is the vCard-style contact model; CreateContact stores it in a contacts folder
// and returns its id. The example creates the contact in a uniquely named contacts
// folder, reads it back and deletes the folder at the end.

using System;
using Aspose.Email.Clients.Exchange;
using Aspose.Email.PersonalInfo;

namespace Aspose.Email.Examples.EWS
{
    internal static class AddContactsToExchangeServerUsingEWS
    {
        public static void Run()
        {
            var folderName = "Aspose-" + Guid.NewGuid().ToString("N").Substring(0, 8);

            using (var client = ClientBuilder.Ews(AuthType.ModernWithAppPermission))
            {
                var folder = client.CreateFolder(client.MailboxInfo.ContactsUri, folderName, ExchangeFolderType.Contact);

                try
                {
                    var contact = new Contact
                    {
                        Gender = Gender.Male,
                        DisplayName = "Frank Lin",
                        CompanyName = "ABC Co.",
                        JobTitle = "Executive Manager"
                    };
                    contact.PhoneNumbers.Add(new PhoneNumber
                    {
                        Number = "+1 555 0100",
                        Category = PhoneNumberCategory.Home
                    });
                    contact.AssociatedPersons.Add(new AssociatedPerson
                    {
                        Name = "Catherine",
                        Category = AssociatedPersonCategory.Spouse
                    });
                    contact.Urls.Add(new Url
                    {
                        Href = "https://www.example.com/frank",
                        Category = UrlCategory.HomePage
                    });
                    contact.EmailAddresses.Add(new EmailAddress
                    {
                        Address = "frank.lin@example.com",
                        DisplayName = "Frank Lin",
                        Category = EmailAddressCategory.Email1
                    });

                    client.CreateContact(folder.Uri, contact);

                    foreach (var stored in client.GetContacts(folder.Uri))
                    {
                        Console.WriteLine($"Created {stored.DisplayName}, {stored.JobTitle} at {stored.CompanyName}");
                        Console.WriteLine($"  phones: {stored.PhoneNumbers.Count}, " +
                                          $"e-mails: {stored.EmailAddresses.Count}, " +
                                          $"web sites: {stored.Urls.Count}");
                    }
                }
                finally
                {
                    client.DeleteFolder(folder.Uri, true);
                }
            }
        }
    }
}

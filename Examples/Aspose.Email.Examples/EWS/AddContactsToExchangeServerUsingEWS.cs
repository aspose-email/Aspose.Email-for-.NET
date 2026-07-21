using Aspose.Email.Clients.Exchange.WebService;
using Aspose.Email.Mime;
using Aspose.Email.PersonalInfo;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;

namespace Aspose.Email.Examples.EWS
{
    class AddContactsToExchangeServerUsingEWS
    {
        public static void Run()
        {
            var client = ClientBuilder.Ews(AuthType.ModernWithAppPermission);

            //Create New Contact
            var contact = new Contact()
            {
                Gender = Gender.Male,
                DisplayName = "Frank Lin",
                CompanyName = "ABC Co.",
                JobTitle = "Executive Manager"
            };

            //Add Phone numbers
            contact.PhoneNumbers.Add(new PhoneNumber { Number = "123456789", Category = PhoneNumberCategory.Home });

            //contact's associated persons
            contact.AssociatedPersons.Add(new AssociatedPerson { Name = "Catherine", Category = AssociatedPersonCategory.Spouse });
            contact.AssociatedPersons.Add(new AssociatedPerson { Name = "Bob", Category = AssociatedPersonCategory.Child });
            contact.AssociatedPersons.Add(new AssociatedPerson { Name = "Merry", Category = AssociatedPersonCategory.Sister });

            //URLs
            contact.Urls.Add(new Url { Href = "www.blog.com", Category = UrlCategory.Blog });
            contact.Urls.Add(new Url { Href = "www.homepage.com", Category = UrlCategory.HomePage });

            //Set contact's Email address
            contact.EmailAddresses.Add(new EmailAddress { Address = "Frank.Lin@Abc.com", DisplayName = "Frank Lin", Category = EmailAddressCategory.Email1 });

            try
            {
                client.CreateContact(contact);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
    }
}

using Aspose.Email.Clients.Exchange.WebService;
using Aspose.Email.Mime;
using Aspose.Email.PersonalInfo;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/

namespace Aspose.Email.Examples.CSharp.Email.Exchange_EWS
{
    class AddContactsToExchangeServerUsingEWS
    {
        public static void Run()
        {
            // ExStart:AddContactsToExchangeServerUsingEWS
            // Set mailboxURI, Username, password, domain information
            string mailboxUri = "https://ex2010/ews/exchange.asmx";
            string username = "test.exchange";
            string password = "pwd";
            string domain = "ex2010.local";
            NetworkCredential credentials = new NetworkCredential(username, password, domain);
            IEWSClient client = EWSClient.GetEWSClient(mailboxUri, credentials);

            //Create New Contact
            Contact contact = new Contact();

            //Set general info
            contact.Gender = Gender.Male;
            contact.DisplayName = "Frank Lin";
            contact.CompanyName = "ABC Co.";
            contact.JobTitle = "Executive Manager";

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
            // ExEnd:AddContactsToExchangeServerUsingEWS
        }
    }
}

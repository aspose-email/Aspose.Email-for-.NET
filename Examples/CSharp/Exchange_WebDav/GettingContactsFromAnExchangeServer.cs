using Aspose.Email.Clients.Exchange.Dav;
using Aspose.Email.Mime;
using Aspose.Email.Mapi;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using Aspose.Email.PersonalInfo;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/

namespace Aspose.Email.Examples.CSharp.Email.Exchange_WebDav
{
    class GettingContactsFromAnExchangeServer
    {
        public static void Run()
        {
            try
            {
                // ExStart:GettingContactsFromAnExchangeServer
                string mailboxURI = "http://ex2003/exchange/administrator"; // WebDAV
                string username = "administrator";
                string password = "pwd";
                string domain = "domain.local";

                // Credentials for connecting to Exchange Server
                NetworkCredential credential = new NetworkCredential(username, password, domain);
                ExchangeClient client = new ExchangeClient(mailboxURI, credential);

                // List all the contacts
                Contact[] contacts = client.GetContacts(client.MailboxInfo.ContactsUri);
                foreach (MapiContact contact in contacts)
                {
                    // Display name and email address
                    Console.WriteLine("Name: " + contact.NameInfo.DisplayName + ", Email Address: " + contact.ElectronicAddresses.Email1);
                }
                // ExEnd:GettingContactsFromAnExchangeServer
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
    }
}

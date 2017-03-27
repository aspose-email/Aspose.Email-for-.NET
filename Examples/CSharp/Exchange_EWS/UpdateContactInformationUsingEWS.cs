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
    class UpdateContactInformationUsingEWS
    {
        public static void Run()
        {
            try
            {
                string mailboxUri = "https://ex2010/ews/exchange.asmx";
                string username = "test.exchange";
                string password = "pwd";
                string domain = "ex2010.local";
                NetworkCredential credentials = new NetworkCredential(username, password, domain);
            
                // ExStart:UpdateContactInformationUsingEWS
                IEWSClient client = EWSClient.GetEWSClient(mailboxUri, credentials);
                
                // List all the contacts and Loop through all contacts
                Contact[] contacts = client.GetContacts(client.MailboxInfo.ContactsUri);
                Contact contact = contacts[0];
                Console.WriteLine("Name: " + contact.DisplayName);
                contact.DisplayName = "David Ch";
                client.UpdateContact(contact);
                // ExEnd:UpdateContactInformationUsingEWS
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
    }
}

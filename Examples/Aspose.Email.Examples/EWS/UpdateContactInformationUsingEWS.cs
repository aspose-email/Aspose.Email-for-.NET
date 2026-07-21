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
            
                IEWSClient client = EWSClient.GetEWSClient(mailboxUri, credentials);
                
                // List all the contacts and Loop through all contacts
                Contact[] contacts = client.GetContacts(client.MailboxInfo.ContactsUri);
                Contact contact = contacts[0];
                Console.WriteLine("Name: " + contact.DisplayName);
                contact.DisplayName = "David Ch";
                client.UpdateContact(contact);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
    }
}

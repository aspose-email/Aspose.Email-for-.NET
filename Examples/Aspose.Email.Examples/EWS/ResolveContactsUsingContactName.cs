using Aspose.Email.Clients.Exchange.WebService;
using Aspose.Email.Mime;
using Aspose.Email.Mapi;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Aspose.Email.PersonalInfo;

namespace Aspose.Email.Examples.EWS
{
    class ResolveContactsUsingContactName
    {
        public static void Run()
        {
            try
            {
                // Create instance of IEWSClient class by giving credentials
                IEWSClient client = EWSClient.GetEWSClient("https://outlook.office365.com/ews/exchange.asmx", "testUser", "pwd", "domain");

                // List all the contacts
                Contact[] contacts = client.ResolveContacts("Changed Name", ExchangeListContactsOptions.FetchPhoto);
                foreach (MapiContact contact in contacts)
                {
                    // Display name and email address
                    Console.WriteLine("Name: " + contact.NameInfo.DisplayName + ", Email Address: " + contact.ElectronicAddresses.Email1);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
    }
}

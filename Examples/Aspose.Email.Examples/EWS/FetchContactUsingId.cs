using Aspose.Email.Clients.Exchange.WebService;
using Aspose.Email.Mime;
using Aspose.Email.PersonalInfo;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Aspose.Email.Examples.EWS
{
    class FetchContactUsingId
    {
        public static void Run()
        {
            try
            {
                // Create instance of IEWSClient class by giving credentials
                IEWSClient client = EWSClient.GetEWSClient("https://exchange.aspose.com/ews/exchange.asmx", "asposeemail.test3", "Aspose2016", "");
                string id = client.GetContacts(client.MailboxInfo.ContactsUri)[0].Id.EWSId;

                Contact fetchedContact = client.GetContact(id);

                Console.WriteLine("Name: " + fetchedContact.DisplayName);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
    }
}

using Aspose.Email.Clients.Exchange.WebService;
using Aspose.Email.Mime;
using Aspose.Email.PersonalInfo;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Aspose.Email.Examples.EWS
{
    class DeleteContactsFromExchangeServerUsingEWS
    {
        public static void Run()
        {
            try
            {
                // Create instance of EWSClient class by giving credentials
                IEWSClient client = EWSClient.GetEWSClient("https://outlook.office365.com/ews/exchange.asmx", "testUser", "pwd", "domain");

                string strContactToDelete = "John Teddy";
                Contact[] contacts = client.GetContacts(client.MailboxInfo.ContactsUri);
                foreach (Contact contact in contacts)
                {
                    if (contact.DisplayName.Equals(strContactToDelete))
                        client.DeleteItem(contact.Id.EWSId, DeletionOptions.DeletePermanently);
                    
                }
                client.Dispose();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
    }
}

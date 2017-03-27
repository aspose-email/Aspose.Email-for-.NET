using Aspose.Email.Clients.Exchange.WebService;
using Aspose.Email.Mime;
using Aspose.Email.Mapi;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Aspose.Email.PersonalInfo;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/

namespace Aspose.Email.Examples.CSharp.Email.Exchange_EWS
{
    class GettingContactsUsingEWS
    {
        public static void Run()
        {
            try
            {
                // ExStart:GettingContactsUsingEWS
                // Create instance of IEWSClient class by giving credentials
                IEWSClient client = EWSClient.GetEWSClient("https://outlook.office365.com/ews/exchange.asmx", "testUser", "pwd", "domain");

                // List all the contacts
                Contact[] contacts = client.GetContacts(client.MailboxInfo.ContactsUri);
                foreach (MapiContact contact in contacts)
                {
                    // Display name and email address
                    Console.WriteLine("Name: " + contact.NameInfo.DisplayName + ", Email Address: " + contact.ElectronicAddresses.Email1);
                }
                // ExEnd:GettingContactsUsingEWS
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
    }
}

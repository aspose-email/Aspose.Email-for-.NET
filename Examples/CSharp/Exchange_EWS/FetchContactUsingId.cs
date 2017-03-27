using Aspose.Email.Clients.Exchange.WebService;
using Aspose.Email.Mime;
using Aspose.Email.PersonalInfo;
using System;
using System.Collections.Generic;
using System.Linq;
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
    class FetchContactUsingId
    {
        public static void Run()
        {
            try
            {
                // Create instance of IEWSClient class by giving credentials
                IEWSClient client = EWSClient.GetEWSClient("https://exchange.aspose.com/ews/exchange.asmx", "asposeemail.test3", "Aspose2016", "");
                string id = client.GetContacts(client.MailboxInfo.ContactsUri)[0].Id.EWSId;

                // ExStart:FetchContactUsingId

                Contact fetchedContact = client.GetContact(id);
                // ExEnd:FetchContactUsingId

                Console.WriteLine("Name: " + fetchedContact.DisplayName);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
    }
}

using Aspose.Email.Clients.Exchange.WebService;
using Aspose.Email.Mime;
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
    class ExpandPublicDistributionList
    {
        public static void Run()
        {
            try
            {
                // Set mailboxURI, Username, password, domain information
                string mailboxUri = "https://ex2010/ews/exchange.asmx";
                string username = "test.exchange";
                string password = "pwd";
                string domain = "ex2010.local";
                NetworkCredential credentials = new NetworkCredential(username, password, domain);
                IEWSClient client = EWSClient.GetEWSClient(mailboxUri, credentials);

                // ExStart:ExpandPublicDistributionList

                MailAddressCollection members = client.ExpandDistributionList(new MailAddress("public.distribution.list@host.com"));
                foreach (MailAddress member in members)
                {
                    Console.WriteLine(member.Address);
                }
                // ExEnd:ExpandPublicDistributionList
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
    }
}

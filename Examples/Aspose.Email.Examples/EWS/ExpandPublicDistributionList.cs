using Aspose.Email.Clients.Exchange.WebService;
using Aspose.Email.Mime;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;

namespace Aspose.Email.Examples.EWS
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

                MailAddressCollection members = client.ExpandDistributionList(new MailAddress("public.distribution.list@host.com"));
                foreach (MailAddress member in members)
                {
                    Console.WriteLine(member.Address);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
    }
}

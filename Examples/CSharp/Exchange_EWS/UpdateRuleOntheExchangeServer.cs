using System;
using System.Net;
using Aspose.Email.Clients.Exchange.WebService;
using Aspose.Email.Mime;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/

namespace Aspose.Email.Examples.CSharp.Email.Exchange_EWS
{
    class UpdateRuleOntheExchangeServer
    {
        public static void Run()
        {

            try
            {

                // ExStart:UpdateRuleOntheExchangeServer           
                // Set mailboxURI, Username, password, domain information
                string mailboxURI = "https://ex2010/ews/exchange.asmx";
                string username = "test.exchange";
                string password = "pwd";
                string domain = "ex2010.local";

                // Connect to the Exchange Server
                NetworkCredential credential = new NetworkCredential(username, password, domain);
                IEWSClient client = EWSClient.GetEWSClient(mailboxURI, credential);

                Console.WriteLine("Connected to Exchange server");

                // Get all Inbox Rules
                InboxRule[] inboxRules = client.GetInboxRules();

                // Loop through each rule
                foreach (InboxRule inboxRule in inboxRules)
                {
                    Console.WriteLine("Display Name: " + inboxRule.DisplayName);
                    if (inboxRule.DisplayName == "Message from client ABC")
                    {
                        Console.WriteLine("Updating the rule....");
                        inboxRule.Conditions.FromAddresses[0] = new MailAddress("administrator@ex2010.local", true);
                        client.UpdateInboxRule(inboxRule);
                    }
                }
                // ExEnd:UpdateRuleOntheExchangeServer
            }
            catch (Exception ex)
            {

                Console.Write(ex.Message);
            }
        }
    }
}
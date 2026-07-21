using System;
using System.Net;
using Aspose.Email.Clients.Exchange;
using Aspose.Email.Clients.Exchange.WebService;

namespace Aspose.Email.Examples.EWS
{
    class UpdateRuleOntheExchangeServer
    {
        public static void Run()
        {
            try
            {
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
            }
            catch (Exception ex)
            {
                Console.Write(ex.Message);
            }
        }
    }
}
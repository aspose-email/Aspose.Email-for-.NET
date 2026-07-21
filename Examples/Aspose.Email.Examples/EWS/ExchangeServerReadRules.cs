using System;
using System.Net;
using Aspose.Email.Clients.Exchange;
using Aspose.Email.Clients.Exchange.WebService;

namespace Aspose.Email.Examples.EWS
{
    class ExchangeServerReadRules
    {
        public static void Run()
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

            // Display information about each rule
            foreach (InboxRule inboxRule in inboxRules)
            {
                Console.WriteLine("Display Name: " + inboxRule.DisplayName);

                // Check if there is a "From Address" condition
                if (inboxRule.Conditions.FromAddresses.Count > 0)
                {
                    foreach (MailAddress fromAddress in inboxRule.Conditions.FromAddresses)
                    {
                        Console.WriteLine("From: " + fromAddress.DisplayName + " - " + fromAddress.Address);
                    }
                }
                // Check if there is a "Subject Contains" condition
                if (inboxRule.Conditions.ContainsSubjectStrings.Count > 0)
                {
                    foreach (String subject in inboxRule.Conditions.ContainsSubjectStrings)
                    {
                        Console.WriteLine("Subject contains: " + subject);
                    }
                }
                // Check if there is a "Move to Folder" action
                if (inboxRule.Actions.MoveToFolder.Length > 0)
                {
                    Console.WriteLine("Move message to folder: " + inboxRule.Actions.MoveToFolder);
                }
            }
        }        
    }
}
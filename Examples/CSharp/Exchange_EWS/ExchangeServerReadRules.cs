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
    class ExchangeServerReadRules
    {
        public static void Run()
        {
            // ExStart:ExchangeServerReadRules
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
            // ExEnd:ExchangeServerReadRules
        }        
    }
}
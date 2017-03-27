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
    class CreateNewRuleOntheExchangeServer
    {
        public static void Run()
        {
            // ExStart:CreateNewRuleOntheExchangeServer
            // Set Exchange Server 2010 web service URL, Username, password, domain information
            string mailboxURI = "https://ex2010/ews/exchange.asmx";
            string username = "test.exchange";
            string password = "pwd";
            string domain = "ex2010.local";

            // Connect to the Exchange Server
            NetworkCredential credential = new NetworkCredential(username, password, domain);
            IEWSClient client = EWSClient.GetEWSClient(mailboxURI, credential);

            Console.WriteLine("Connected to Exchange server");

            InboxRule rule = new InboxRule();
            rule.DisplayName = "Message from client ABC";

            // Add conditions
            RulePredicates newRules = new RulePredicates();
            // Set Subject contains string "ABC" and Add the conditions
            newRules.ContainsSubjectStrings.Add("ABC");
            newRules.FromAddresses.Add(new MailAddress("administrator@ex2010.local", true));
            rule.Conditions = newRules;

            // Add Actions and Move the message to a folder
            RuleActions newActions = new RuleActions();
            newActions.MoveToFolder = "120:AAMkADFjMjNjMmNjLWE3NzgtNGIzNC05OGIyLTAwNTgzNjRhN2EzNgAuAAAAAABbwP+Tkhs0TKx1GMf0D/cPAQD2lptUqri0QqRtJVHwOKJDAAACL5KNAAA=AQAAAA==";
            rule.Actions = newActions;
            client.CreateInboxRule(rule);
            // ExEnd:CreateNewRuleOntheExchangeServer
        }        
    }
}
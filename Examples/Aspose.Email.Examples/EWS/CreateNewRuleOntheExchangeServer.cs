using System;
using Aspose.Email.Clients.Exchange;

namespace Aspose.Email.Examples.EWS
{
    class CreateNewRuleOntheExchangeServer
    {
        public static void Run()
        {
            using (var client = ClientBuilder.Ews(AuthType.ModernWithAppPermission))
            {
                Console.WriteLine("Connected to Exchange server");

                // Add conditions
                var newRules = new RulePredicates();

                // Set Subject contains string "ABC" and Add the conditions
                newRules.ContainsSubjectStrings.Add("ABC");
                newRules.FromAddresses.Add(new MailAddress("administrator@ex2010.local", true));

                // Add Actions and Move the message to a folder
                var newActions = new RuleActions
                {
                    MoveToFolder = client.MailboxInfo.DeletedItemsUri
                };

                var rule = new InboxRule
                {
                    DisplayName = "Message from client ABC",
                    Conditions = newRules,
                    Actions = newActions
                };

                client.CreateInboxRule(rule);
            }
        }        
    }
}
// Demonstrates creating an Inbox rule that the server applies to incoming mail.
//
// A rule combines conditions (RulePredicates: words in the subject, senders, ...),
// optional exceptions and actions (RuleActions: move, copy, delete, forward, ...).
// Exchange runs it on the server, so it works even when no client is open. The example
// creates a rule with a unique name and deletes it again at the end.

using System;
using System.Linq;
using Aspose.Email.Clients.Exchange;

namespace Aspose.Email.Examples.EWS
{
    internal static class CreateNewRuleOntheExchangeServer
    {
        public static void Run()
        {
            var ruleName = "Aspose example rule " + Guid.NewGuid().ToString("N").Substring(0, 8);

            using (var client = ClientBuilder.Ews(AuthType.ModernWithAppPermission))
            {
                var conditions = new RulePredicates();
                conditions.ContainsSubjectStrings.Add("[newsletter]");
                conditions.FromAddresses.Add(new MailAddress("news@example.com", true));

                var rule = new InboxRule
                {
                    DisplayName = ruleName,
                    Conditions = conditions,
                    Actions = new RuleActions { MoveToFolder = client.MailboxInfo.DeletedItemsUri }
                };

                client.CreateInboxRule(rule);

                var created = client.GetInboxRules().FirstOrDefault(r => r.DisplayName == ruleName);
                Console.WriteLine(created != null
                    ? $"Created '{created.DisplayName}', enabled: {created.IsEnabled}"
                    : "The new rule was not found.");

                if (created != null)
                    client.DeleteInboxRule(created.RuleId);
            }
        }
    }
}

// Demonstrates deleting an Inbox rule.
//
// Rules are deleted by id. CreateInboxRule does not return one, so the example creates
// a rule with a unique name, finds it again in GetInboxRules to read its RuleId, and
// then removes it with DeleteInboxRule.

using System;
using System.Linq;
using Aspose.Email.Clients.Exchange;

namespace Aspose.Email.Examples.EWS
{
    internal static class DeleteInboxRuleUsingEWS
    {
        public static void Run()
        {
            var ruleName = "Aspose example rule " + Guid.NewGuid().ToString("N").Substring(0, 8);

            using (var client = ClientBuilder.Ews(AuthType.ModernWithAppPermission))
            {
                var rule = InboxRule.CreateRuleMoveContaining(new[] { "[newsletter]" },
                    client.MailboxInfo.DeletedItemsUri);
                rule.DisplayName = ruleName;
                client.CreateInboxRule(rule);
                Console.WriteLine($"Created rule '{ruleName}'.");

                var created = client.GetInboxRules().FirstOrDefault(r => r.DisplayName == ruleName);
                if (created == null)
                {
                    Console.WriteLine("The new rule was not found.");
                    return;
                }

                client.DeleteInboxRule(created.RuleId);
                Console.WriteLine($"Deleted rule {created.RuleId}.");

                var stillThere = client.GetInboxRules().Any(r => r.DisplayName == ruleName);
                Console.WriteLine($"Rule still present: {stillThere}");
            }
        }
    }
}

// Demonstrates changing an existing Inbox rule.
//
// Read the rule with GetInboxRules, change its conditions, exceptions or actions, and
// pass it to UpdateInboxRule. The example creates a rule with a unique name, adds a
// second sender to it, and deletes it again at the end.

using System;
using System.Linq;
using Aspose.Email.Clients.Exchange;

namespace Aspose.Email.Examples.EWS
{
    internal static class UpdateRuleOntheExchangeServer
    {
        public static void Run()
        {
            var ruleName = "Aspose example rule " + Guid.NewGuid().ToString("N").Substring(0, 8);

            using (var client = ClientBuilder.Ews(AuthType.ModernWithAppPermission))
            {
                var newRule = InboxRule.CreateRuleMoveFrom(new MailAddress("news@example.com", true),
                    client.MailboxInfo.DeletedItemsUri);
                newRule.DisplayName = ruleName;
                client.CreateInboxRule(newRule);

                // Find it by its unique name, so no other rule is touched.
                var rule = client.GetInboxRules().FirstOrDefault(r => r.DisplayName == ruleName);
                if (rule == null)
                {
                    Console.WriteLine("The new rule was not found.");
                    return;
                }

                try
                {
                    rule.Conditions.FromAddresses.Add(new MailAddress("offers@example.com", true));
                    client.UpdateInboxRule(rule);

                    var updated = client.GetInboxRules().First(r => r.RuleId == rule.RuleId);
                    Console.WriteLine($"'{updated.DisplayName}' now applies to:");
                    foreach (var from in updated.Conditions.FromAddresses)
                        Console.WriteLine($"  {from.Address}");
                }
                finally
                {
                    client.DeleteInboxRule(rule.RuleId);
                }
            }
        }
    }
}

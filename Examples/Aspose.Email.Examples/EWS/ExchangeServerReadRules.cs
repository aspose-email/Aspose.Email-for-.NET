// Demonstrates reading the Inbox rules of a mailbox.
//
// GetInboxRules returns every rule with its name, order and state, its conditions -
// senders, subject words and more - and its actions, such as the folder messages are
// moved to.

using System;

namespace Aspose.Email.Examples.EWS
{
    internal static class ExchangeServerReadRules
    {
        public static void Run()
        {
            using (var client = ClientBuilder.Ews(AuthType.ModernWithAppPermission))
            {
                var rules = client.GetInboxRules();
                Console.WriteLine($"{rules.Length} Inbox rule(s):");

                foreach (var rule in rules)
                {
                    Console.WriteLine($"\n  {rule.DisplayName} (priority {rule.Priority}, enabled: {rule.IsEnabled})");

                    if (rule.Conditions != null)
                    {
                        foreach (var from in rule.Conditions.FromAddresses)
                            Console.WriteLine($"    from:             {from.Address}");
                        foreach (var subject in rule.Conditions.ContainsSubjectStrings)
                            Console.WriteLine($"    subject contains: {subject}");
                    }

                    if (!string.IsNullOrEmpty(rule.Actions?.MoveToFolder))
                        Console.WriteLine($"    moves to:         {rule.Actions.MoveToFolder}");
                }
            }
        }
    }
}

// Demonstrates the two ways Graph sorts incoming mail: inbox rules, which act on the
// messages themselves, and Focused Inbox overrides, which decide whether a given
// sender lands in Focused or Other.

using System;
using System.Linq;
using Aspose.Email;
using Aspose.Email.Clients.Exchange;
using Aspose.Email.Clients.Graph;

namespace Aspose.Email.Examples.Graph
{
    internal static class ManageGraphInboxRules
    {
        public static void Run()
        {
            if (!ClientBuilder.IsGraphConfigured)
            {
                GraphExampleInfo.PrintNotConfigured();
                return;
            }

            using (var client = ClientBuilder.Graph(AuthType.ModernWithAppPermission))
            {
                var rules = client.ListRules(null);
                Console.WriteLine($"{rules.Count} inbox rule(s):");

                foreach (var existing in rules)
                    Console.WriteLine($"  {existing.DisplayName} (enabled: {existing.IsEnabled})");

                // InboxRule offers factories for the common shapes; the Conditions and
                // Actions properties are there when a rule needs more than one criterion.
                var archive = client.GetFolder(KnownFolders.Archive);
                var rule = InboxRule.CreateRuleMoveFrom(new MailAddress("client@example.com"), archive.ItemId);
                rule.DisplayName = "Move messages from a client";
                rule.IsEnabled = true;

                var created = client.CreateRule(rule);
                Console.WriteLine($"\nCreated rule: {created.DisplayName}");

                var fetched = client.FetchRule(created.RuleId);
                Console.WriteLine($"Fetched rule: {fetched.DisplayName}, enabled: {fetched.IsEnabled}");

                fetched.IsEnabled = false;
                var updated = client.UpdateRule(fetched);
                Console.WriteLine($"Disabled it: enabled is now {updated.IsEnabled}");

                // Focused Inbox overrides are managed separately from the rules.
                var overrides = client.ListOverrides(null);
                Console.WriteLine($"\n{overrides.Count} classification override(s):");

                foreach (var item in overrides.Take(10))
                    Console.WriteLine($"  {item.Sender.Address} -> {item.ClassifyAs}");

                var newsletter = new MailAddress("newsletter@example.com");
                var createdOverride = client.CreateOrUpdateOverride(newsletter, ClassificationType.Other);
                Console.WriteLine($"Always classify {createdOverride.Sender.Address} as {createdOverride.ClassifyAs}");

                createdOverride.ClassifyAs = ClassificationType.Focused;
                var updatedOverride = client.UpdateOverride(createdOverride);
                Console.WriteLine($"Changed it to {updatedOverride.ClassifyAs}");
            }
        }
    }
}

// Demonstrates the Outlook categories of a mailbox.
//
// A message stores only the category name; the colour lives in the mailbox-wide list
// that these calls manage. CategoryPreset is Outlook's fixed palette - preset1 is the
// red one, preset2 orange, and so on.

using System;
using System.Linq;
using Aspose.Email.Clients.Graph;

namespace Aspose.Email.Examples.Graph
{
    internal static class ManageGraphCategories
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
                var categories = client.ListCategories(null);
                Console.WriteLine($"{categories.Count} category/categories defined:");

                foreach (var category in categories)
                    Console.WriteLine($"  {category.DisplayName}: {category.Preset} ({category.Color})");

                var created = client.CreateCategory("Aspose.Email", CategoryPreset.Preset5);
                Console.WriteLine($"\nCreated: {created.DisplayName} -> {created.Preset} ({created.Color})");

                var fetched = client.FetchCategory(created.Id);
                Console.WriteLine($"Fetched: {fetched.DisplayName}");

                // Only the colour can be changed; the name identifies the category.
                fetched.Preset = CategoryPreset.Preset9;
                var updated = client.UpdateCategory(fetched);
                Console.WriteLine($"Recoloured to: {updated.Preset} ({updated.Color})");

                client.Delete(created.Id);
                Console.WriteLine("Deleted the category.");

                var remaining = client.ListCategories(null).Count(c => c.DisplayName == "Aspose.Email");
                Console.WriteLine($"Categories left with that name: {remaining}");
            }
        }
    }
}

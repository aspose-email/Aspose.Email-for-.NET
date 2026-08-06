// Demonstrates how to read the category list of an OLM storage. Messages carry only
// the category name, so the colour is looked up in the storage-wide list.

using System;
using System.Linq;
using Aspose.Email.Storage.Olm;

namespace Aspose.Email.Examples.OLM
{
    internal static class GetCategoryColorsFromOLM
    {
        public static void Run()
        {
            using (var olm = new OlmStorage(Data.Mapi/"SampleOLM.olm"))
            {
                var categories = olm.GetCategories();
                Console.WriteLine($"Categories defined in the storage: {categories.Count}");

                foreach (var category in categories)
                {
                    // The colour is a hexadecimal value in #rrggbb form.
                    Console.WriteLine($"  {category.Name} -> {category.Color}");
                }

                Console.WriteLine();

                foreach (var folder in olm.FolderHierarchy)
                {
                    if (!folder.HasMessages)
                        continue;

                    foreach (var msg in olm.EnumerateMessages(folder))
                    {
                        if (msg.Categories == null || msg.Categories.Length == 0)
                            continue;

                        Console.WriteLine($"{msg.Subject}");
                        foreach (var msgCategory in msg.Categories)
                        {
                            var match = categories.FirstOrDefault(
                                c => c.Name.Equals(msgCategory, StringComparison.OrdinalIgnoreCase));

                            Console.WriteLine(match != null
                                ? $"  {msgCategory} -> {match.Color}"
                                : $"  {msgCategory} -> colour not defined in this storage");
                        }
                    }
                }
            }
        }
    }
}

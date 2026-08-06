// Demonstrates how to read the category list of a PST or OST. A message stores only
// the category name, so the colour has to be looked up in the storage-wide list.

using System;
using Aspose.Email.Mapi;
using Aspose.Email.Storage.Pst;

namespace Aspose.Email.Examples.MAPI
{
    internal static class GetCategoryColorsFromPST
    {
        public static void Run()
        {
            // The category list is a property of the storage, not of the messages, so it
            // is only populated in files where Outlook itself defined the categories.
            using (var pst = PersonalStorage.FromFile(Data.Mapi/"SampleOstFile.ost", false))
            {
                var availableCategories = pst.GetCategories();

                Console.WriteLine($"Categories defined in the storage: {availableCategories.Count}");
                foreach (var category in availableCategories)
                    Console.WriteLine($"  {category.Name} -> {category.Color}");

                Console.WriteLine("\nMessages carrying a category:");
                if (Walk(pst, pst.RootFolder, availableCategories) == 0)
                    Console.WriteLine("  (none in this storage)");
            }
        }

        // Returns how many messages carried at least one category.
        private static int Walk(PersonalStorage pst, FolderInfo folder,
            System.Collections.Generic.List<PstItemCategory> availableCategories)
        {
            var tagged = 0;

            foreach (var messageInfo in folder.EnumerateMessages())
            {
                var messageCategoryList = FollowUpManager.GetCategories(pst.ExtractMessage(messageInfo));
                if (messageCategoryList.Count == 0)
                    continue;

                Console.WriteLine($"  {messageInfo.Subject}");
                tagged++;

                foreach (var messageCategory in messageCategoryList)
                {
                    var category = availableCategories.Find(
                        c => c.Name.Equals(messageCategory, StringComparison.OrdinalIgnoreCase));

                    Console.WriteLine(category != null
                        ? $"    {category.Name} -> {category.Color}"
                        : $"    {messageCategory} -> colour not defined in this storage");
                }
            }

            foreach (var subFolder in folder.GetSubFolders())
                tagged += Walk(pst, subFolder, availableCategories);

            return tagged;
        }
    }
}

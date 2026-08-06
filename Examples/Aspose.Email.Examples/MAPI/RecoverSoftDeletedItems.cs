// Demonstrates how to find items that were soft-deleted from a PST or OST - removed
// from their folder by Outlook but not yet purged from the file - and write them back
// out to disk.
//
// Two APIs cover this: FindAndExtractSoftDeletedItems materialises the whole set at
// once, FindAndEnumerateSoftDeletedItems streams it, which matters for large storages.
// Neither creates soft-deleted items; how many are found depends on the input file.

using System;
using System.IO;
using Aspose.Email.Storage.Pst;

namespace Aspose.Email.Examples.MAPI
{
    internal static class RecoverSoftDeletedItems
    {
        public static void Run()
        {
            var outputDir = Data.OutSub("RecoveredItems");

            using (var pst = PersonalStorage.FromFile(Data.Mapi/"SampleOstFile.ost", false))
            {
                // Streaming variant: never holds the whole set in memory.
                var found = 0;
                foreach (var entry in pst.FindAndEnumerateSoftDeletedItems())
                {
                    Console.WriteLine($"Subject: {entry.Item.Subject}");
                    Console.WriteLine($"Deleted from folder id: {entry.FolderId}");
                    found++;
                }

                Console.WriteLine($"{found} soft-deleted item(s) found by enumeration.\n");

                // Batch variant: the recovered items are saved under the folder they were
                // deleted from.
                var entries = pst.FindAndExtractSoftDeletedItems();
                Console.WriteLine($"{entries.Count} soft-deleted item(s) found by extraction.");

                for (var index = 0; index < entries.Count; index++)
                {
                    var folderInfo = pst.GetFolderById(entries[index].FolderId);
                    var folderDir = Path.Combine(outputDir, folderInfo.DisplayName);
                    Directory.CreateDirectory(folderDir);

                    var msg = entries[index].Item;
                    msg.Save(Path.Combine(folderDir, $"{index}.msg"));

                    Console.WriteLine($"  recovered '{msg.Subject}' from {folderInfo.DisplayName}");
                }

                if (entries.Count == 0)
                    Console.WriteLine("  (this sample storage has none - run it against a file where Outlook deleted items)");
            }

            Console.WriteLine($"\nRecovered items would be written to {outputDir}");
        }
    }
}

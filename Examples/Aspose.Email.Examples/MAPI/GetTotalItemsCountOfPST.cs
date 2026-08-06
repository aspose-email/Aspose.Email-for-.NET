// Demonstrates how to get the number of items a PST holds in one call, instead of
// walking the folder tree and adding up the per-folder counts.

using System;
using Aspose.Email.Storage.Pst;

namespace Aspose.Email.Examples.MAPI
{
    internal static class GetTotalItemsCountOfPST
    {
        public static void Run()
        {
            using (var pst = PersonalStorage.FromFile(Data.Mapi/"Sub.pst", false))
            {
                // The count comes from the message store, so it covers every folder.
                Console.WriteLine($"Total items in the storage: {pst.Store.GetTotalItemsCount()}");

                foreach (var folder in pst.RootFolder.GetSubFolders())
                    Console.WriteLine($"  {folder.DisplayName}: {folder.ContentCount} item(s)");
            }
        }
    }
}

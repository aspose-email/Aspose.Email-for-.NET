// Demonstrates how to get the number of items an OLM storage holds in one call,
// instead of walking the folder tree and adding up the per-folder counts.

using System;
using Aspose.Email.Storage.Olm;

namespace Aspose.Email.Examples.OLM
{
    internal static class GetTotalItemsCountOfOLM
    {
        public static void Run()
        {
            using (var olm = new OlmStorage(Data.Mapi/"SampleOLM.olm"))
            {
                Console.WriteLine($"Total items in the storage: {olm.GetTotalItemsCount()}");

                foreach (var folder in olm.FolderHierarchy)
                    Console.WriteLine($"  {folder.Name}: {folder.MessageCount} message(s)");
            }
        }
    }
}

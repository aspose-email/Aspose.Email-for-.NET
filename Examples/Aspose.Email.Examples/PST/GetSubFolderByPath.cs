// Demonstrates how to reach a nested PST folder by its path in a single call,
// rather than calling GetSubFolder once per level.

using System;
using Aspose.Email.Storage.Pst;

namespace Aspose.Email.Examples.PST
{
    internal static class GetSubFolderByPath
    {
        public static void Run()
        {
            var pstPath = Data.Out/"GetSubFolderByPath_out.pst";

            using (var pst = PersonalStorage.Create(pstPath, FileFormatVersion.Unicode))
            {
                pst.RootFolder.AddSubFolder(@"Inbox\Reports\Jan", true);

                // The third argument is what turns the backslashes into a path lookup;
                // the second one makes the name comparison case-insensitive.
                var folder = pst.RootFolder.GetSubFolder(@"Inbox\Reports\Jan", true, true);

                Console.WriteLine($"Found by path: {folder.DisplayName}");
                Console.WriteLine($"Entry id:      {folder.EntryIdString}");

                // Without the path flag the same string is treated as a single folder name.
                var notFound = pst.RootFolder.GetSubFolder(@"Inbox\Reports\Jan", true, false);
                Console.WriteLine($"Same name without path handling: {(notFound == null ? "not found" : notFound.DisplayName)}");
            }

            Console.WriteLine($"\nSaved to {pstPath}");
        }
    }
}

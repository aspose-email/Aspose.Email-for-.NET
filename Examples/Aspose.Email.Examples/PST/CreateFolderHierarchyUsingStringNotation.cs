// Demonstrates how to create a whole folder hierarchy in one call, by passing a
// path instead of adding each folder separately.

using System;
using Aspose.Email.Storage.Pst;

namespace Aspose.Email.Examples.PST
{
    internal static class CreateFolderHierarchyUsingStringNotation
    {
        public static void Run()
        {
            var pstPath = Data.Out/"CreateFolderHierarchyUsingStringNotation.pst";

            using (var personalStorage = PersonalStorage.Create(pstPath, FileFormatVersion.Unicode))
            {
                // The second argument makes the missing parents be created as well, so
                // Inbox and Folder1 do not have to exist beforehand.
                personalStorage.RootFolder.AddSubFolder(@"Inbox\Folder1\Folder2", true);

                Console.WriteLine(@"Created folder hierarchy Inbox\Folder1\Folder2");
                Console.WriteLine($"Saved to {pstPath}");
            }
        }
    }
}

// Demonstrates how to merge a single folder from one PST into another PST,
// tracking the moved items through the ItemMoved event.

using System;
using System.IO;
using Aspose.Email.Storage.Pst;

namespace Aspose.Email.Examples.MAPI
{
    internal static class MergeFolderFromAnotherPstFile
    {
        private const string MergedFolderName = "FolderFromAnotherPst";

        private static int totalAdded;

        public static void Run()
        {
            totalAdded = 0;

            // Work on a copy: the merge writes into the destination PST, and examples
            // must never modify the shared input data.
            var destinationPath = Data.Out/"MergeFolderFromAnotherPstFile_out.pst";
            File.Copy(Data.Mapi/"destination.pst", destinationPath, true);

            using (var destinationPst = PersonalStorage.FromFile(destinationPath))
            using (var sourcePst = PersonalStorage.FromFile(Data.Mapi/"source.pst", false))
            {
                // The sample destination.pst already contains this folder, and AddSubFolder
                // throws if the name is taken - so reuse the folder when it is already there.
                var destinationFolder = destinationPst.RootFolder.GetSubFolder(MergedFolderName)
                                        ?? destinationPst.RootFolder.AddSubFolder(MergedFolderName);

                var sourceFolder = sourcePst.GetPredefinedFolder(StandardIpmFolder.Inbox);

                destinationFolder.ItemMoved += OnItemMoved;
                destinationFolder.MergeWith(sourceFolder);

                Console.WriteLine($"Merged \"{sourceFolder.DisplayName}\" into \"{MergedFolderName}\".");
                Console.WriteLine($"Total messages added: {totalAdded}");
                Console.WriteLine($"Saved to {destinationPath}");
            }
        }

        private static void OnItemMoved(object sender, ItemMovedEventArgs e)
        {
            totalAdded++;
        }
    }
}

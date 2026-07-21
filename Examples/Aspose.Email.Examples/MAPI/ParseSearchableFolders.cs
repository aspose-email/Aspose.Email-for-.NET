// Demonstrates how to walk a PST folder hierarchy including search folders and print each folder's path.

using System;
using System.IO;
using Aspose.Email.Storage.Pst;

namespace Aspose.Email.Examples.MAPI
{
    internal static class ParseSearchableFolders
    {
        public static void Run()
        {
            string path = Data.Mapi/"PersonalStorage.pst";
            using (Stream s = File.OpenRead(path))
            {
                PersonalStorage pst = PersonalStorage.FromStream(s);
                WalkFolders(pst.RootFolder, "N/A");
            }
        }

        private static void WalkFolders(FolderInfo folder, string parentFolderName)
        {
            string displayName = string.IsNullOrEmpty(folder.DisplayName) ? "ROOT" : folder.DisplayName;
            Console.WriteLine("DisplayName = {0}; Parent.DisplayName = {1}", displayName, parentFolderName);

            if (!folder.HasSubFolders)
                return;

            foreach (FolderInfo subfolder in folder.GetSubFolders(FolderKind.Search | FolderKind.Normal))
                WalkFolders(subfolder, displayName);
        }
    }
}

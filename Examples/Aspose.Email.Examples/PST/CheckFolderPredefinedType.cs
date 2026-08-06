// Demonstrates how to tell which PST folders are standard Outlook folders (Inbox,
// Contacts, ...) and which are ordinary user folders.

using System;
using Aspose.Email.Storage.Pst;

namespace Aspose.Email.Examples.PST
{
    internal static class CheckFolderPredefinedType
    {
        public static void Run()
        {
            using (var pst = PersonalStorage.FromFile(Data.Mapi/"PersonalStorage.pst", false))
            {
                CheckFolders(pst.RootFolder.GetSubFolders());
            }
        }

        private static void CheckFolders(FolderInfoCollection folders, string indent = "")
        {
            foreach (var folder in folders)
            {
                Console.Write($"{indent}{folder.DisplayName}: ");

                // false: report only the folder's own role.
                var folderType = folder.GetPredefinedType(false);

                if (folderType != StandardIpmFolder.Unspecified)
                {
                    Console.WriteLine($"standard folder ({folderType})");
                }
                else
                {
                    // true: walk up to the top-level parent, which tells whether this is
                    // a user folder nested under a standard one.
                    folderType = folder.GetPredefinedType(true);

                    Console.WriteLine(folderType != StandardIpmFolder.Unspecified
                        ? $"user folder under a standard parent ({folderType})"
                        : "user folder");
                }

                CheckFolders(folder.GetSubFolders(), indent + "    ");
            }
        }
    }
}

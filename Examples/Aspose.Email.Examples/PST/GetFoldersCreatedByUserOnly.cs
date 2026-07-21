// Demonstrates how to list only the folders a user created, filtering out the
// predefined ones Outlook creates itself (Inbox, Sent Items, and so on).

using System;
using Aspose.Email.Storage.Pst;
using Aspose.Email.Tools.Search;

namespace Aspose.Email.Examples.PST
{
    internal static class GetFoldersCreatedByUserOnly
    {
        public static void Run()
        {
            // The source is only read, so open it read-only.
            using (var pst = PersonalStorage.FromFile(Data.Mapi/"Outlook.pst", false))
            {
                var queryBuilder = new PersonalStorageQueryBuilder();
                queryBuilder.OnlyFoldersCreatedByUser.Equals(true);

                var subfolders = pst.RootFolder.GetSubFolders(queryBuilder.GetQuery());
                foreach (var folder in subfolders)
                    Console.WriteLine(folder.DisplayName);

                Console.WriteLine($"\n{subfolders.Count} of {pst.RootFolder.GetSubFolders().Count} " +
                                  "folder(s) were created by the user.");
            }
        }
    }
}

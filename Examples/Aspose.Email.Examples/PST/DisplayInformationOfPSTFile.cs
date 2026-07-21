// Demonstrates how to display the name and the item counts of every folder in
// an Outlook PST file.

using System;
using Aspose.Email.Storage.Pst;

namespace Aspose.Email.Examples.PST
{
    internal static class DisplayInformationOfPSTFile
    {
        public static void Run()
        {
            // The source is only read, so open it read-only.
            using (var personalStorage = PersonalStorage.FromFile(Data.Mapi/"PersonalStorage.pst", false))
            {
                var folderInfoCollection = personalStorage.RootFolder.GetSubFolders();

                foreach (var folderInfo in folderInfoCollection)
                {
                    Console.WriteLine($"Folder:             {folderInfo.DisplayName}");
                    Console.WriteLine($"Total items:        {folderInfo.ContentCount}");
                    Console.WriteLine($"Total unread items: {folderInfo.ContentUnreadCount}");
                    Console.WriteLine("-----------------------------------");
                }

                Console.WriteLine($"{folderInfoCollection.Count} folder(s) in total.");
            }
        }
    }
}

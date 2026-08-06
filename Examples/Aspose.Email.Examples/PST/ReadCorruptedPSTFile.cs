// Demonstrates how to salvage a damaged PST or OST: FindMessages and FindSubfolders
// walk the storage by entry id, so a folder whose own record is unreadable does not
// stop the rest of the file from being read.

using System;
using Aspose.Email.Storage.Pst;

namespace Aspose.Email.Examples.PST
{
    internal static class ReadCorruptedPSTFile
    {
        public static void Run()
        {
            using (var pst = PersonalStorage.FromFile(Data.Mapi/"PersonalStorage.pst", false))
            {
                Explore(pst, pst.RootFolder.EntryIdString);
            }
        }

        private static void Explore(PersonalStorage pst, string rootFolderId, string indent = "")
        {
            foreach (var messageId in pst.FindMessages(rootFolderId))
            {
                try
                {
                    var msg = pst.ExtractMessage(messageId);
                    Console.WriteLine($"{indent}- {msg.Subject}");
                }
                catch
                {
                    // A single unreadable message must not abort the whole scan.
                    Console.WriteLine($"{indent}- message reading error. Entry id: {messageId}");
                }
            }

            foreach (var subFolderId in pst.FindSubfolders(rootFolderId))
            {
                if (subFolderId == rootFolderId)
                    continue;

                try
                {
                    var subfolder = pst.GetFolderById(subFolderId);
                    Console.WriteLine($"{indent}{subfolder.DisplayName}");
                }
                catch
                {
                    Console.WriteLine($"{indent}folder reading error. Entry id: {subFolderId}");
                }

                Explore(pst, subFolderId, indent + "    ");
            }
        }
    }
}

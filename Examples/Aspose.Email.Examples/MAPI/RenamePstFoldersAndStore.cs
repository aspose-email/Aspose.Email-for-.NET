// Demonstrates renaming the parts of a PST in place: a single message property, a
// folder's display name, and the name of the message store itself - the label Outlook
// shows at the top of the folder tree.

using System;
using System.IO;
using System.Linq;
using System.Text;
using Aspose.Email.Mapi;
using Aspose.Email.Storage.Pst;

namespace Aspose.Email.Examples.MAPI
{
    internal static class RenamePstFoldersAndStore
    {
        public static void Run()
        {
            // Work on a copy: everything here writes into the storage.
            var pstPath = Data.Out/"RenamePstFoldersAndStore_out.pst";
            File.Copy(Data.Mapi/"Sub.pst", pstPath, true);

            using (var pst = PersonalStorage.FromFile(pstPath))
            {
                var inbox = pst.RootFolder.GetSubFolder("Inbox");
                var messageInfo = inbox.GetContents()[0];
                Console.WriteLine($"Message before: {messageInfo.Subject}");

                // ChangeMessage rewrites properties of one message without extracting,
                // modifying and re-adding it.
                var updated = new MapiPropertyCollection();
                updated.Add(MapiPropertyTag.PR_SUBJECT_W,
                    new MapiProperty(MapiPropertyTag.PR_SUBJECT_W, Encoding.Unicode.GetBytes("Renamed subject")));

                pst.ChangeMessage(messageInfo.EntryIdString, updated);
                Console.WriteLine($"Message after:  {pst.RootFolder.GetSubFolder("Inbox").GetContents()[0].Subject}");

                Console.WriteLine($"\nStore before:  {pst.Store.DisplayName}");
                pst.Store.ChangeDisplayName("Renamed store");
                Console.WriteLine($"Store after:   {pst.Store.DisplayName}");

                Console.WriteLine("\nFolders before: " + Names(pst));
                inbox.ChangeDisplayName("Renamed inbox");
                Console.WriteLine("Folders after:  " + Names(pst));
            }

            Console.WriteLine($"\nSaved to {pstPath}");
        }

        private static string Names(PersonalStorage pst)
        {
            return string.Join(", ", pst.RootFolder.GetSubFolders().Select(f => f.DisplayName).Take(6)) + ", ...";
        }
    }
}

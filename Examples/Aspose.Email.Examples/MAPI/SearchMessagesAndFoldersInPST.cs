// Demonstrates the range of criteria PersonalStorageQueryBuilder supports when
// searching a PST - by importance, class, flags, size, and by folder name.

using System;
using Aspose.Email.Mapi;
using Aspose.Email.Storage.Pst;
using Aspose.Email.Tools.Search;

namespace Aspose.Email.Examples.MAPI
{
    internal static class SearchMessagesAndFoldersInPST
    {
        public static void Run()
        {
            // The source is only read, so open it read-only.
            using (var personalStorage = PersonalStorage.FromFile(Data.Mapi/"Outlook.pst", false))
            {
                var folder = personalStorage.RootFolder.GetSubFolder("Inbox");

                Report("High importance", folder, b => b.Importance.Equals((int)MapiImportance.High));

                Report("Message class IPM.Note", folder, b => b.MessageClass.Equals("IPM.Note"));

                // Criteria set on the same builder are combined with AND.
                Report("High importance with attachments", folder, b =>
                {
                    b.Importance.Equals((int)MapiImportance.High);
                    b.HasFlags(MapiMessageFlags.MSGFLAG_HASATTACH);
                });

                Report("Larger than 15 KB", folder, b => b.MessageSize.Greater(15000));

                Report("Unread", folder, b => b.HasNoFlags(MapiMessageFlags.MSGFLAG_READ));

                Report("Unread with attachments", folder, b =>
                {
                    b.HasNoFlags(MapiMessageFlags.MSGFLAG_READ);
                    b.HasFlags(MapiMessageFlags.MSGFLAG_HASATTACH);
                });

                // The same builder also searches folders rather than messages.
                ReportFolders("Named 'SubInbox'", folder, b => b.FolderName.Equals("SubInbox"));
                ReportFolders("Having subfolders", folder, b => b.HasSubfolders());
            }
        }

        private static void Report(string label, FolderInfo folder, Action<PersonalStorageQueryBuilder> configure)
        {
            var builder = new PersonalStorageQueryBuilder();
            configure(builder);
            Console.WriteLine($"{label,-34}: {folder.GetContents(builder.GetQuery()).Count} message(s)");
        }

        private static void ReportFolders(string label, FolderInfo folder, Action<PersonalStorageQueryBuilder> configure)
        {
            var builder = new PersonalStorageQueryBuilder();
            configure(builder);
            Console.WriteLine($"{label,-34}: {folder.GetSubFolders(builder.GetQuery()).Count} folder(s)");
        }
    }
}

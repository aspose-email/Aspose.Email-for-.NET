// Demonstrates reading parts of a PST message without materialising the whole thing.
// ExtractProperty fetches a single MAPI property by tag and ExtractAttachments fetches
// only the attachments, both by entry id - far cheaper than ExtractMessage when a scan
// needs one field per message.

using System;
using System.Linq;
using Aspose.Email.Mapi;
using Aspose.Email.Storage.Pst;

namespace Aspose.Email.Examples.MAPI
{
    internal static class ExtractPropertyWithoutMessage
    {
        public static void Run()
        {
            using (var pst = PersonalStorage.FromFile(Data.Mapi/"Sub.pst", false))
            {
                var inbox = pst.RootFolder.GetSubFolder("Inbox");

                // The storage can enumerate a folder straight from its entry id, so a
                // FolderInfo does not have to be kept around.
                var folderId = inbox.EntryIdString;
                Console.WriteLine($"Messages in the folder: {pst.EnumerateMessages(folderId).Count()}");

                Console.WriteLine("\nFirst three, by entry id:");
                foreach (var messageInfo in pst.EnumerateMessages(folderId, 0, 3))
                {
                    // One property, without extracting the message.
                    var subject = pst.ExtractProperty(messageInfo.EntryId, MapiPropertyTag.PR_SUBJECT_W);
                    var sender = pst.ExtractProperty(messageInfo.EntryId, MapiPropertyTag.PR_SENDER_NAME_W);

                    Console.WriteLine($"  subject: {Show(subject)}");
                    Console.WriteLine($"  sender:  {Show(sender)}");

                    // Attachments only - the body and the rest of the properties are
                    // never read.
                    var attachments = pst.ExtractAttachments(messageInfo.EntryIdString);
                    Console.WriteLine($"  attachments: {attachments.Count}");

                    foreach (var attachment in attachments)
                        Console.WriteLine($"    {attachment.LongFileName}");
                }
            }
        }

        private static string Show(MapiProperty property)
        {
            return property == null ? "(not set)" : property.GetString();
        }
    }
}

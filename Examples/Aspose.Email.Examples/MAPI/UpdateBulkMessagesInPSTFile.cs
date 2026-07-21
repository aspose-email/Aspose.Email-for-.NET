// Demonstrates how to update the properties of many PST messages in a single bulk
// operation, instead of extracting, changing and re-saving them one by one.

using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Email.Mapi;
using Aspose.Email.Storage.Pst;
using Aspose.Email.Tools.Search;

namespace Aspose.Email.Examples.MAPI
{
    internal static class UpdateBulkMessagesInPSTFile
    {
        public static void Run()
        {
            // Work on a copy: the update rewrites messages, and examples must never
            // modify the shared input data.
            var pstPath = Data.Out/"UpdateBulkMessagesInPSTFile_out.pst";
            File.Copy(Data.Mapi/"Sub.pst", pstPath, true);

            using (var personalStorage = PersonalStorage.FromFile(pstPath))
            {
                var inbox = personalStorage.RootFolder.GetSubFolder("Inbox");

                // Replace this with your own criteria - this one simply matches messages
                // that the sample PST is known to contain.
                var queryBuilder = new PersonalStorageQueryBuilder();
                queryBuilder.Subject.Contains("Test Message");

                var changeList = new List<string>();
                foreach (var messageInfo in inbox.GetContents(queryBuilder.GetQuery()))
                    changeList.Add(messageInfo.EntryIdString);

                // The properties to apply to every matching message.
                var updatedProperties = new MapiPropertyCollection
                {
                    {
                        MapiPropertyTag.PR_SUBJECT_W,
                        new MapiProperty(MapiPropertyTag.PR_SUBJECT_W, Encoding.Unicode.GetBytes("New Subject"))
                    },
                    {
                        MapiPropertyTag.PR_IMPORTANCE,
                        new MapiProperty(MapiPropertyTag.PR_IMPORTANCE, BitConverter.GetBytes((long)2))
                    }
                };

                // One call for the whole batch.
                inbox.ChangeMessages(changeList, updatedProperties);

                Console.WriteLine($"Updated {changeList.Count} message(s) in {pstPath}");
            }
        }
    }
}

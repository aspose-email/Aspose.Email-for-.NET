// Demonstrates how to extract non-MSG attachments from all messages in a PST inbox folder.

using System;
using Aspose.Email.Storage.Pst;

namespace Aspose.Email.Examples.MAPI
{
    internal static class ExtractAttachmentsFromPstMessages
    {
        public static void Run()
        {
            var outputDir = Data.OutSub("Attachments");

            using (var pst = PersonalStorage.FromFile(Data.Mapi/"Outlook.pst"))
            {
                var folder = pst.RootFolder.GetSubFolder("Inbox");
                var extracted = 0;

                foreach (var entryId in folder.EnumerateMessagesEntryId())
                {
                    foreach (var attachment in pst.ExtractAttachments(entryId))
                    {
                        // Embedded messages are skipped: this example is about file
                        // attachments, and a nested MSG needs different handling.
                        if (!string.IsNullOrEmpty(attachment.LongFileName) &&
                            !attachment.LongFileName.Contains(".msg"))
                        {
                            attachment.Save(outputDir/attachment.LongFileName);
                            Console.WriteLine($"Extracted: {attachment.LongFileName}");
                            extracted++;
                        }
                    }
                }

                Console.WriteLine($"\nExtracted {extracted} attachment(s) to {outputDir}");
            }
        }
    }
}

// Demonstrates how to write a message straight from a PST into a stream, without
// materialising it as a MapiMessage first.

using System;
using System.IO;
using Aspose.Email.Storage.Pst;

namespace Aspose.Email.Examples.MAPI
{
    internal static class SaveMessagesDirectlyFromPSTToStream
    {
        public static void Run()
        {
            var outputDir = Data.OutSub("PstMessages");

            // The source is only read, so open it read-only.
            using (var personalStorage = PersonalStorage.FromFile(Data.Mapi/"Outlook.pst", false))
            {
                var inbox = personalStorage.RootFolder.GetSubFolder("Inbox");
                var toMemory = 0;
                var toFile = 0;

                // Into a memory stream, addressed by the message's entry id.
                foreach (var messageInfo in inbox.EnumerateMessages())
                {
                    using (var memoryStream = new MemoryStream())
                    {
                        personalStorage.SaveMessageToStream(messageInfo.EntryIdString, memoryStream);
                        toMemory++;
                    }
                }

                // Or straight into a file stream.
                foreach (var messageInfo in inbox.EnumerateMessages())
                {
                    var path = outputDir/(ToFileName(messageInfo.Subject, toFile) + ".msg");
                    using (var fs = File.Create(path))
                    {
                        personalStorage.SaveMessageToStream(messageInfo.EntryIdString, fs);
                        toFile++;
                    }
                }

                // EnumerateMessagesEntryId() is the cheaper option when only the id is
                // needed, as it does not build a MessageInfo per message.
                var byEntryId = 0;
                foreach (var entryId in inbox.EnumerateMessagesEntryId())
                {
                    using (var ms = new MemoryStream())
                    {
                        personalStorage.SaveMessageToStream(entryId, ms);
                        byEntryId++;
                    }
                }

                Console.WriteLine($"Saved {toMemory} message(s) to memory streams.");
                Console.WriteLine($"Saved {toFile} message(s) to {outputDir}");
                Console.WriteLine($"Saved {byEntryId} message(s) enumerated by entry id.");
            }
        }

        // A subject may contain characters that are not valid in a file name.
        private static string ToFileName(string subject, int index)
        {
            if (string.IsNullOrWhiteSpace(subject))
                return $"message-{index}";

            foreach (var invalid in Path.GetInvalidFileNameChars())
                subject = subject.Replace(invalid, ' ');

            return subject.Trim();
        }
    }
}

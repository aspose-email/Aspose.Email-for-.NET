// Demonstrates the two-pass pattern over an OLM storage: list the summary properties
// first, decide from them what is worth reading, then pull those messages in full by
// their entry id.
//
// The ids have to be collected in one complete pass. A folder reads as a forward-only
// cursor, so interleaving "read one id" and "extract that message" would advance the
// cursor between the two and hand back the wrong item.

using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Email.Storage.Olm;

namespace Aspose.Email.Examples.OLM
{
    internal static class ExtractMessagesFromOLMById
    {
        public static void Run()
        {
            var outputDir = Data.OutSub("OlmMessages");

            using (var storage = new OlmStorage(Data.Mapi/"SampleOLM.olm"))
            {
                var folder = FindLargestFolder(storage.GetFolders());
                if (folder == null)
                {
                    Console.WriteLine("No folder in this storage holds any messages.");
                    return;
                }

                Console.WriteLine($"Folder: {folder.Path}\n");

                // One complete pass, so the cursor ends up rewound.
                var infos = folder.EnumerateMessages().ToList();

                foreach (var info in infos)
                {
                    Console.WriteLine($"{info.Subject}");
                    Console.WriteLine($"  class:       {info.MessageClass}");
                    Console.WriteLine($"  sent:        {info.Date:yyyy-MM-dd HH:mm}");
                    Console.WriteLine($"  modified:    {info.ModifiedDate:yyyy-MM-dd HH:mm}");
                    Console.WriteLine($"  from:        {Address(info.From)}");
                    Console.WriteLine($"  to:          {Recipients(info.To)}");
                    Console.WriteLine($"  attachments: {info.HasAttachments}");
                }

                Console.WriteLine("\nExtracting the ones that carry attachments:");
                var saved = 0;

                foreach (var info in infos.Where(i => i.HasAttachments))
                {
                    // ExtractMapiMessage also takes the OlmMessageInfo directly.
                    var message = storage.ExtractMapiMessage(info.EntryId);

                    var outputPath = outputDir/$"message-{saved}.msg";
                    message.Save(outputPath);

                    Console.WriteLine($"  {message.Subject} -> {message.Attachments.Count} attachment(s)");
                    saved++;
                }

                Console.WriteLine($"\nSaved {saved} message(s) to {outputDir}");
            }
        }

        private static string Address(Aspose.Email.Mapi.MapiElectronicAddress address)
        {
            return address == null || string.IsNullOrEmpty(address.EmailAddress) ? "(none)" : address.EmailAddress;
        }

        private static string Recipients(List<Aspose.Email.Mapi.MapiElectronicAddress> addresses)
        {
            if (addresses == null || addresses.Count == 0)
                return "(none)";

            return string.Join(", ", addresses.Select(a => a.EmailAddress));
        }

        private static OlmFolder FindLargestFolder(List<OlmFolder> folders)
        {
            OlmFolder best = null;

            foreach (var folder in folders)
            {
                if (best == null || folder.MessageCount > best.MessageCount)
                    best = folder;

                if (folder.SubFolders != null && folder.SubFolders.Count > 0)
                {
                    var deeper = FindLargestFolder(folder.SubFolders);
                    if (deeper != null && (best == null || deeper.MessageCount > best.MessageCount))
                        best = deeper;
                }
            }

            return best != null && best.MessageCount > 0 ? best : null;
        }
    }
}

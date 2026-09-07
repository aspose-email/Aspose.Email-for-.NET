// Demonstrates the fault-tolerant way to get a message out of a PST.
//
// ExtractMessage throws when a message is damaged, which aborts a bulk export.
// TryToSaveMessage writes whatever it can and reports the outcome instead: Success,
// PartiallySaved with the properties it had to skip, or Corrupted.

using System;
using System.IO;
using Aspose.Email.Storage.Pst;

namespace Aspose.Email.Examples.MAPI
{
    internal static class SaveMessageSafelyFromPST
    {
        public static void Run()
        {
            var outputDir = Data.OutSub("SafeExport");

            using (var pst = PersonalStorage.FromFile(Data.Mapi/"Sub.pst", false))
            {
                // TryToGetFolderById is the non-throwing counterpart of GetFolderById.
                FolderInfo inbox;
                if (!pst.TryToGetFolderById(pst.RootFolder.GetSubFolder("Inbox").EntryIdString, out inbox))
                {
                    Console.WriteLine("Inbox not found.");
                    return;
                }

                var index = 0;
                var success = 0;
                var partial = 0;
                var corrupted = 0;

                foreach (var messageInfo in pst.EnumerateMessages(inbox.EntryIdString, 0, 5))
                {
                    var outputPath = outputDir/$"message-{index}.msg";

                    using (var stream = File.Create(outputPath))
                    {
                        var result = pst.TryToSaveMessage(messageInfo.EntryIdString, stream);

                        Console.WriteLine($"{index}. {messageInfo.Subject}");
                        Console.WriteLine($"   status:      {result.Status}");
                        Console.WriteLine($"   attachments: {result.Attachments.Count}");

                        if (result.MissedProperties.Count > 0)
                        {
                            Console.WriteLine($"   skipped {result.MissedProperties.Count} property/properties:");
                            foreach (var property in result.MissedProperties)
                                Console.WriteLine($"     {property}");
                        }

                        switch (result.Status)
                        {
                            case SaveStatus.Success: success++; break;
                            case SaveStatus.PartiallySaved: partial++; break;
                            default: corrupted++; break;
                        }
                    }

                    index++;
                }

                Console.WriteLine($"\n{success} saved cleanly, {partial} partially, {corrupted} corrupted.");
                Console.WriteLine($"Written to {outputDir}");
            }
        }
    }
}

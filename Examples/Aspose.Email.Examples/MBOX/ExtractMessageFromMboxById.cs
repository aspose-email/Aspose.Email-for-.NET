// Demonstrates how to pull a single message out of an mbox storage by its entry id,
// so a listing pass and a fetch pass can be kept apart.

using System;
using Aspose.Email.Storage.Mbox;

namespace Aspose.Email.Examples.MBOX
{
    internal static class ExtractMessageFromMboxById
    {
        public static void Run()
        {
            var outputDir = Data.OutSub("MboxById");

            using (var reader = MboxStorageReader.CreateReader(Data.Mbox/"ExampleMbox.mbox", new MboxLoadOptions()))
            {
                var index = 0;

                foreach (var messageInfo in reader.EnumerateMessageInfo())
                {
                    // The entry id stays valid for the lifetime of the reader.
                    var eml = reader.ExtractMessage(messageInfo.EntryId, new EmlLoadOptions());

                    var outputPath = outputDir/$"message-{index}.eml";
                    eml.Save(outputPath, SaveOptions.DefaultEml);

                    Console.WriteLine($"{messageInfo.EntryId} -> {eml.Subject}");
                    index++;
                }

                Console.WriteLine($"\nSaved {index} message(s) to {outputDir}");
            }
        }
    }
}

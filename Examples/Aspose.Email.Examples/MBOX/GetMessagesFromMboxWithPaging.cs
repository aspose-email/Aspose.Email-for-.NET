// Demonstrates how to read an mbox storage in fixed-size batches, so a large file
// never has to be walked from the start for every page.

using System;
using Aspose.Email.Storage.Mbox;

namespace Aspose.Email.Examples.MBOX
{
    internal static class GetMessagesFromMboxWithPaging
    {
        public static void Run()
        {
            using (var reader = MboxStorageReader.CreateReader(Data.Mbox/"ExampleMbox.mbox", new MboxLoadOptions()))
            {
                var total = reader.GetTotalItemsCount();
                Console.WriteLine($"Messages in the storage: {total}");

                const int pageSize = 2;

                for (var startIndex = 0; startIndex < total; startIndex += pageSize)
                {
                    Console.WriteLine($"--- messages {startIndex}..{startIndex + pageSize - 1} ---");

                    foreach (var message in reader.EnumerateMessages(startIndex, pageSize))
                        Console.WriteLine($"  {message.Subject}");
                }
            }
        }
    }
}

// Demonstrates how to read an mbox storage in fixed-size batches.
//
// The catch: a reader is a forward-only cursor over the stream, so a single reader
// cannot serve more than one page - the second call would carry on from wherever the
// first one stopped instead of seeking back. Each page therefore gets its own reader.

using System;
using Aspose.Email.Storage.Mbox;

namespace Aspose.Email.Examples.MBOX
{
    internal static class GetMessagesFromMboxWithPaging
    {
        public static void Run()
        {
            // The shipped sample holds a single message, which is too few to page
            // through, so build a storage with enough messages to show the batching.
            var mboxPath = BuildStorage(7);

            int total;
            using (var reader = MboxStorageReader.CreateReader(mboxPath, new MboxLoadOptions()))
                total = reader.GetTotalItemsCount();

            Console.WriteLine($"Messages in the storage: {total}");

            const int pageSize = 3;

            for (var startIndex = 0; startIndex < total; startIndex += pageSize)
            {
                Console.WriteLine($"--- messages {startIndex}..{Math.Min(startIndex + pageSize, total) - 1} ---");

                // A fresh reader per page: reusing one would return the wrong slice.
                using (var reader = MboxStorageReader.CreateReader(mboxPath, new MboxLoadOptions()))
                {
                    foreach (var message in reader.EnumerateMessages(startIndex, pageSize))
                        Console.WriteLine($"  {message.Subject}");
                }
            }
        }

        private static string BuildStorage(int messageCount)
        {
            var mboxPath = Data.Out/"GetMessagesFromMboxWithPaging_in.mbox";

            using (var writer = new MboxrdStorageWriter(mboxPath, new MboxSaveOptions()))
            {
                for (var i = 1; i <= messageCount; i++)
                {
                    writer.WriteMessage(new MailMessage(
                        "sender@domain.com", "receiver@domain.com", $"Message {i}", $"Body of message {i}"));
                }
            }

            return mboxPath;
        }
    }
}

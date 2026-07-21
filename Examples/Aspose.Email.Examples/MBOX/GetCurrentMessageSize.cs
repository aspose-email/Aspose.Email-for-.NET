// Demonstrates how to track how much of an mbox storage has been consumed while
// reading it, using MboxrdStorageReader.CurrentDataSize.

using System;
using System.IO;
using Aspose.Email.Storage.Mbox;

namespace Aspose.Email.Examples.MBOX
{
    internal static class GetCurrentMessageSize
    {
        public static void Run()
        {
            using (var stream = new FileStream(Data.Mbox/"ExampleMbox.mbox", FileMode.Open, FileAccess.Read))
            using (var reader = new MboxrdStorageReader(stream, new MboxLoadOptions()))
            {
                var count = 0;
                long previousDataSize = 0;

                MailMessage msg;
                while ((msg = reader.ReadNextMessage()) != null)
                {
                    using (msg)
                    {
                        count++;

                        // CurrentDataSize is the running total of bytes read so far, so the
                        // size of this message is how much it advanced.
                        var messageSize = reader.CurrentDataSize - previousDataSize;
                        previousDataSize = reader.CurrentDataSize;

                        Console.WriteLine($"{count,3}. {messageSize,8:N0} bytes  {msg.Subject}");
                    }
                }

                Console.WriteLine($"\nRead {count} message(s), {reader.CurrentDataSize:N0} bytes in total.");
            }
        }
    }
}

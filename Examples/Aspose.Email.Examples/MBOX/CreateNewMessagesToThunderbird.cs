// Demonstrates how to append a new message to a Thunderbird mbox storage file.

using System;
using System.IO;
using Aspose.Email.Storage.Mbox;

namespace Aspose.Email.Examples.MBOX
{
    internal static class CreateNewMessagesToThunderbird
    {
        public static void Run()
        {
            // Work on a copy: writing appends to the mbox, and examples must never modify
            // the shared input data that the other Thunderbird examples read.
            var mboxPath = Data.Out/"CreateNewMessagesToThunderbird_out.mbox";
            File.Copy(Data.Mbox/"ExampleMbox.mbox", mboxPath, true);

            using (var stream = new FileStream(mboxPath, FileMode.Open, FileAccess.ReadWrite))
            using (var writer = new MboxrdStorageWriter(stream, new MboxSaveOptions()))
            {
                var message = new MailMessage(
                    "from@domain.com",
                    "to@domain.com",
                    "Message added by Aspose.Email",
                    "This message was appended to an existing mbox storage.")
                {
                    IsDraft = false
                };

                // The writer appends to the end of the existing storage.
                writer.WriteMessage(message);
            }

            Console.WriteLine($"Appended a new message to {mboxPath}");
        }
    }
}

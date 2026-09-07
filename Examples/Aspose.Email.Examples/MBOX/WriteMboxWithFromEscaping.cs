// Demonstrates the "From " escaping rule of the mboxrd format.
//
// A line starting with "From " is what separates one message from the next, so a body
// line that happens to begin the same way would split the message in two. Setting
// MboxSaveOptions.FromShouldBeEscaped writes such lines as ">From " instead.

using System;
using System.IO;
using System.Linq;
using Aspose.Email.Storage.Mbox;

namespace Aspose.Email.Examples.MBOX
{
    internal static class WriteMboxWithFromEscaping
    {
        private const string Body = "Line one.\r\nFrom the desk of Alice\r\nLine three.\r\n";

        public static void Run()
        {
            Write("unescaped", false);
            Write("escaped", true);
        }

        private static void Write(string label, bool escapeFrom)
        {
            var mboxPath = Data.Out/$"WriteMboxWithFromEscaping_{label}.mbox";

            var saveOptions = new MboxSaveOptions
            {
                FromShouldBeEscaped = escapeFrom,

                // LeaveOpen matters only for the stream-based constructor: it decides
                // whether disposing the writer also closes the stream underneath it.
                LeaveOpen = false
            };

            using (var writer = new MboxrdStorageWriter(mboxPath, saveOptions))
            {
                var message = new MailMessage(
                    "from@domain.com", "to@domain.com", $"From-line test ({label})", Body);

                // WriteMessage hands back the entry id of the message it just appended,
                // which is what MboxStorageReader.ExtractMessage takes.
                var entryId = writer.WriteMessage(message);

                Console.WriteLine($"FromShouldBeEscaped = {escapeFrom}");
                Console.WriteLine($"  entry id:    {entryId}");
                Console.WriteLine($"  base stream: {writer.BaseStream.GetType().Name}");
            }

            var bodyLine = File.ReadLines(mboxPath).FirstOrDefault(l => l.Contains("the desk of"));
            Console.WriteLine($"  body line on disk: '{bodyLine?.TrimEnd()}'");
            Console.WriteLine($"  saved to {mboxPath}\n");
        }
    }
}

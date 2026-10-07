// Demonstrates controlling the order of headers that share a name.
//
// A message can carry the same header several times - Received, or a custom header
// stamped by each system a message passes. Headers.Add puts a new occurrence after the
// existing ones; Headers.Insert puts it before them, so it becomes the first occurrence,
// which is the one most programs read. This example works on a file and runs offline.

using System;
using System.IO;
using System.Linq;

namespace Aspose.Email.Examples.IMAP
{
    internal static class InsertHeaderAtSpecificLocation
    {
        private const string HeaderName = "secret-header";

        public static void Run()
        {
            // The sample file already carries "secret-header: mystery".
            var message = MailMessage.Load(Data.Imap/"InsertHeaders.eml");

            message.Headers.Add(HeaderName, "added");
            message.Headers.Insert(HeaderName, "inserted");

            var outputPath = Data.Out/"InsertHeaderAtSpecificLocation_out.eml";
            message.Save(outputPath, SaveOptions.DefaultEml);

            // Read the saved file back to show the order the headers ended up in.
            var occurrences = File.ReadLines(outputPath)
                .TakeWhile(line => line.Length > 0)
                .Where(line => line.StartsWith(HeaderName + ":", StringComparison.OrdinalIgnoreCase));

            Console.WriteLine($"'{HeaderName}' occurrences, in order:");
            foreach (var line in occurrences)
                Console.WriteLine("  " + line);

            Console.WriteLine($"\nSaved to {outputPath}");
        }
    }
}

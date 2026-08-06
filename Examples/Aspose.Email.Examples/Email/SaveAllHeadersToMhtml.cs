// Demonstrates how to keep every original header in an MHTML export. By default only
// the headers MHT normally renders survive the conversion.

using System;
using System.IO;

namespace Aspose.Email.Examples.Email
{
    internal static class SaveAllHeadersToMhtml
    {
        public static void Run()
        {
            var eml = MailMessage.Load(Data.Email/"emlWithHeaders.eml");
            Console.WriteLine($"Headers in the source message: {eml.Headers.Count}");

            var defaultPath = Save(eml, "default", false);
            var allPath = Save(eml, "all-headers", true);

            Console.WriteLine($"\nDefault output:     {new FileInfo(defaultPath).Length} byte(s)");
            Console.WriteLine($"All-headers output: {new FileInfo(allPath).Length} byte(s)");
        }

        private static string Save(MailMessage eml, string label, bool saveAllHeaders)
        {
            var options = SaveOptions.DefaultMhtml;
            options.SaveAllHeaders = saveAllHeaders;

            var outputPath = Data.Out/$"SaveAllHeadersToMhtml_{label}.mhtml";
            eml.Save(outputPath, options);

            Console.WriteLine($"{label}: SaveAllHeaders = {saveAllHeaders} -> {outputPath}");
            return outputPath;
        }
    }
}

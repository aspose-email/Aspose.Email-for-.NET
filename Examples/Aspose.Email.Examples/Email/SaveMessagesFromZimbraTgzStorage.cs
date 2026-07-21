// Demonstrates how to export all messages from a Zimbra TGZ archive to a local directory.

using System;
using System.IO;
using Aspose.Email.Storage.Zimbra;

namespace Aspose.Email.Examples.Email
{
    internal static class SaveMessagesFromZimbraTgzStorage
    {
        public static void Run()
        {
            var outputDir = Data.OutSub("Zimbra");

            using (var reader = new TgzReader(Data.Email/"ZimbraSample.tgz"))
            {
                // ExportTo recreates the archive's folder structure under the target
                // directory, so the messages keep their original layout.
                reader.ExportTo(outputDir);
            }

            var exported = Directory.GetFiles(outputDir, "*", SearchOption.AllDirectories).Length;
            Console.WriteLine($"Exported {exported} message(s) to {outputDir}");
        }
    }
}

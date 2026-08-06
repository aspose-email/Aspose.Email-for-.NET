// Demonstrates how to pack a directory of EML files into a Zimbra TGZ backup.
// TgzBackupBuilder is the counterpart of TgzReader: it writes the archive rather
// than reading one.

using System;
using System.IO;
using Aspose.Email.Storage.Zimbra.TgzBackup;

namespace Aspose.Email.Examples.Email
{
    internal static class BuildTgzBackupFromEml
    {
        public static void Run()
        {
            // Collect a few messages into an input directory - the builder takes a
            // folder of EML files, not individual paths.
            var inputDir = Data.OutSub("TgzBackupInput");
            foreach (var fileName in new[] { "Message.eml", "Attachments.eml", "test.eml" })
                File.Copy(Data.Email/fileName, inputDir/fileName, true);

            Console.WriteLine($"Input directory: {inputDir}");
            Console.WriteLine($"Messages to pack: {Directory.GetFiles(inputDir, "*.eml").Length}");

            var archivePath = Data.Out/"BuildTgzBackupFromEml_out.tgz";

            // The third argument is the folder name the messages get inside the archive.
            var result = TgzBackupBuilder.Build(inputDir, archivePath, "Imported");

            Console.WriteLine($"\nMessages written: {result.MessagesWritten}");
            Console.WriteLine($"Errors: {result.Errors.Count}");
            foreach (var error in result.Errors)
                Console.WriteLine($"  {error}");

            Console.WriteLine($"Archive size: {new FileInfo(archivePath).Length} byte(s)");
            Console.WriteLine($"Saved to {archivePath}");
        }
    }
}

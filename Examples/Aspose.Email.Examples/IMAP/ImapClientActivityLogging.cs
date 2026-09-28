// Demonstrates logging the client's activity - every command sent and every response
// received - to a file, with a separate file per day.
//
// EnableLogger switches logging on, LogFileName sets the file, and UseDateInLogFileName
// adds the date to that name, so a long-running service gets one log per day instead of
// a single ever-growing file. ConfigureLoggingInCode shows the same settings with a fixed
// file name.
//
// The log can contain message content and authentication data - keep it private.

using System;
using System.IO;
using Aspose.Email.Clients.Imap;

namespace Aspose.Email.Examples.IMAP
{
    internal static class ImapClientActivityLogging
    {
        public static void Run()
        {
            var logDir = Data.OutSub("ImapLogs");
            var messagesDir = Data.OutSub("ImapClientActivityLogging");

            using (var client = ClientBuilder.Imap(AuthType.ModernWithDelegatedPermission))
            {
                client.LogFileName = logDir/"ImapClientActivity.log";
                client.UseDateInLogFileName = true;
                client.EnableLogger = true;

                // Some activity worth logging: download a few messages.
                client.SelectFolder(ImapFolderInfo.InBox);
                var messages = client.ListMessages(ImapFolderInfo.InBox, ImapListFields.IdOnly, 3);

                foreach (var info in messages)
                    client.SaveMessage(info.UniqueId, messagesDir/(info.UniqueId + ".eml"));

                Console.WriteLine($"Downloaded {messages.Count} message(s) to {messagesDir}");
            }

            var logFiles = Directory.GetFiles(logDir, "*", SearchOption.AllDirectories);

            Console.WriteLine($"\nLog file(s) in {logDir}:");
            if (logFiles.Length == 0)
                Console.WriteLine("  (none)");

            foreach (var file in logFiles)
                Console.WriteLine($"  {Path.GetFileName(file)}  ({new FileInfo(file).Length} bytes)");
        }
    }
}

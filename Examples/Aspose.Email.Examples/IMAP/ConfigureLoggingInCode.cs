// Demonstrates how to record the IMAP conversation to a log file from code.
//
// The log holds every command the client sent and every response it got, which is the
// first thing to look at when a server misbehaves. Logging can also be configured in
// App.config; setting it on the client works the same on .NET Framework and .NET.
//
// The log can contain message content and authentication data - keep it private.

using System;
using System.IO;
using Aspose.Email.Clients.Imap;

namespace Aspose.Email.Examples.IMAP
{
    internal static class ConfigureLoggingInCode
    {
        public static void Run()
        {
            var logPath = Data.Out/"ConfigureLoggingInCode.log";

            using (var client = ClientBuilder.Imap(AuthType.ModernWithDelegatedPermission))
            {
                client.LogFileName = logPath;
                client.UseDateInLogFileName = false;  // true adds the date to the file name
                client.EnableLogger = true;

                client.SelectFolder(ImapFolderInfo.InBox);
                var messages = client.ListMessages(ImapFolderInfo.InBox, ImapListFields.IdOnly, 3);
                Console.WriteLine($"Listed {messages.Count} message(s) with logging on.");

                // Back to the settings the client started with; nothing below is logged.
                client.ResetLogSettings();
                client.Noop();
            }

            if (!File.Exists(logPath))
            {
                Console.WriteLine($"No log was written to {logPath}.");
                return;
            }

            Console.WriteLine($"\nFirst lines of {logPath}:");

            // The logger may still hold the file open, so allow shared access.
            using (var stream = new FileStream(logPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var reader = new StreamReader(stream))
            {
                string line;
                for (var i = 0; i < 15 && (line = reader.ReadLine()) != null; i++)
                    Console.WriteLine("  " + line);
            }
        }
    }
}

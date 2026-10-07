// Demonstrates logging the client's activity - every command sent and every reply
// received - to a file, with a separate file per day.
//
// EnableLogger switches logging on, LogFileName sets the file, and UseDateInLogFileName
// adds the date to that name, so a long-running service gets one log per day instead of
// a single ever-growing file. ConfigureSmtpLoggingInCode shows the same settings with a
// fixed file name.
//
// The log can contain message content and authentication data - keep it private.

using System;
using System.IO;

namespace Aspose.Email.Examples.SMTP
{
    internal static class SMTPClientActivityLogging
    {
        public static void Run()
        {
            if (!ClientBuilder.IsSmtpConfigured)
            {
                SmtpExampleInfo.PrintNotConfigured();
                return;
            }

            var logDir = Data.OutSub("SmtpLogs");

            using (var client = ClientBuilder.Smtp(AuthType.Basic))
            {
                client.LogFileName = logDir/"SmtpClientActivity.log";
                client.UseDateInLogFileName = true;
                client.EnableLogger = true;

                client.Send(new MailMessage(client.Username, client.Username, "Logged message", "Body"));
                Console.WriteLine($"Sent a message to {client.Username} with logging on.");
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

// Demonstrates how to record the SMTP conversation to a log file from code.
//
// The log holds every command the client sent and every reply it got - the first thing
// to look at when a server refuses a message. It is switched off here, but
// UseDateInLogFileName = true adds the date to the file name, so a long-running service
// gets one log per day.
//
// The log can contain message content and authentication data - keep it private.

using System;
using System.IO;

namespace Aspose.Email.Examples.SMTP
{
    internal static class ConfigureSmtpLoggingInCode
    {
        public static void Run()
        {
            if (!ClientBuilder.IsSmtpConfigured)
            {
                SmtpExampleInfo.PrintNotConfigured();
                return;
            }

            var logPath = Data.Out/"ConfigureSmtpLoggingInCode.log";

            using (var client = ClientBuilder.Smtp(AuthType.Basic))
            {
                client.LogFileName = logPath;
                client.UseDateInLogFileName = false;
                client.EnableLogger = true;

                client.Send(new MailMessage(client.Username, client.Username, "Logged message", "Body"));
                Console.WriteLine($"Sent a message to {client.Username} with logging on.");
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

// Demonstrates sending every .eml file in a folder - an outbox on disk, for instance,
// filled by another program.
//
// Each file is loaded, readdressed and sent over the same connection. A file that cannot
// be loaded or sent is reported and skipped, so one bad file does not stop the rest.

using System;
using System.IO;

namespace Aspose.Email.Examples.SMTP
{
    internal static class LoadingEMLFilesFromDisk
    {
        public static void Run()
        {
            if (!ClientBuilder.IsSmtpConfigured)
            {
                SmtpExampleInfo.PrintNotConfigured();
                return;
            }

            var files = Directory.GetFiles(Data.Smtp, "*.eml");
            var sent = 0;

            using (var client = ClientBuilder.Smtp(AuthType.Basic))
            {
                foreach (var file in files)
                {
                    try
                    {
                        var message = MailMessage.Load(file);

                        // The files are addressed to sample recipients; send them to yourself.
                        message.From = client.Username;
                        message.To.Clear();
                        message.CC.Clear();
                        message.Bcc.Clear();
                        message.To.Add(client.Username);

                        client.Send(message);
                        sent++;
                        Console.WriteLine($"  sent    {Path.GetFileName(file)}: {message.Subject}");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"  skipped {Path.GetFileName(file)}: {ex.Message}");
                    }
                }
            }

            Console.WriteLine($"\n{sent} of {files.Length} file(s) sent to yourself.");
        }
    }
}

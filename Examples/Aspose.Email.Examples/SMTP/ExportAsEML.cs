// Demonstrates how to create an HTML message and save it as an .eml file.
//
// EML is the plain MIME format every mail program understands, which makes it the
// natural format for a message that is to be sent later - SendingEMLFilesWithSMTP and
// LoadingEMLFilesFromDisk send such files.

using System;

namespace Aspose.Email.Examples.SMTP
{
    internal static class ExportAsEML
    {
        public static void Run()
        {
            var message = new MailMessage
            {
                From = "from@example.com",
                Subject = "New message created by Aspose.Email for .NET",
                IsBodyHtml = true,
                HtmlBody = "<b>This line is in bold.</b><br/><br/><font color=\"blue\">This line is in blue.</font>"
            };

            message.To.Add("to1@example.com");
            message.To.Add("to2@example.com");
            message.CC.Add("cc1@example.com");
            message.CC.Add("cc2@example.com");

            var outputPath = Data.Out/"ExportAsEML_out.eml";
            message.Save(outputPath, SaveOptions.DefaultEml);

            Console.WriteLine($"Saved to {outputPath}");
            Console.WriteLine($"  to: {message.To}");
            Console.WriteLine($"  cc: {message.CC}");
        }
    }
}

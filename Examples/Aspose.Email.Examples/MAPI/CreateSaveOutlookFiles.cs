// Demonstrates how to create a MailMessage, convert it to MapiMessage, and save as MSG.

using System;
using Aspose.Email.Mapi;

namespace Aspose.Email.Examples.MAPI
{
    internal static class CreateSaveOutlookFiles
    {
        public static void Run()
        {
            var mailMsg = new MailMessage
            {
                From = "from@domain.com",
                Subject = "creating an outlook message file",
                Body = "This message is created by Aspose.Email"
            };
            mailMsg.To.Add("to@domain.com");

            // FromMailMessage turns the MIME message into Outlook's own representation.
            var outlookMsg = MapiMessage.FromMailMessage(mailMsg);

            var outputPath = Data.Out/"message.msg";
            outlookMsg.Save(outputPath);

            Console.WriteLine($"Message: {outlookMsg.Subject}");
            Console.WriteLine($"Saved to {outputPath}");
        }
    }
}

// Demonstrates how to create a MailMessage, convert it to MapiMessage, and save it as MSG.

using System;
using Aspose.Email.Mapi;

namespace Aspose.Email.Examples.MAPI
{
    internal static class CreatingAndSavingOutlookMessages
    {
        public static void Run()
        {
            var mailMsg = new MailMessage
            {
                From = "sender@domain.com",
                To = "receiver@domain.com",
                Subject = "This is test message",
                Body = "This is test body"
            };

            var outlookMsg = MapiMessage.FromMailMessage(mailMsg);

            var outputPath = Data.Out/"CreatingAndSavingOutlookMessages_out.msg";
            outlookMsg.Save(outputPath);

            Console.WriteLine($"Message: {outlookMsg.Subject}");
            Console.WriteLine($"Saved to {outputPath}");
        }
    }
}

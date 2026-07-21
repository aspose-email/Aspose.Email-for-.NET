// Demonstrates how to create a MailMessage with an HTML body, convert it to MapiMessage, and save as MSG.

using System;
using Aspose.Email.Mapi;

namespace Aspose.Email.Examples.MAPI
{
    internal static class CreatingMsgFilesWithRtfBody
    {
        public static void Run()
        {
            var mailMsg = new MailMessage
            {
                From = "from@domain.com",
                To = "to@domain.com",
                Subject = "subject",
                HtmlBody = "<h3>rtf example</h3><p>creating an <b><u>outlook message (msg)</u></b> file using Aspose.Email.</p>"
            };

            // The HTML body is converted to RTF on the way into MSG, which is the format
            // Outlook uses for formatted text.
            var outlookMsg = MapiMessage.FromMailMessage(mailMsg);

            var outputPath = Data.Out/"CreatingMSGFilesWithRTFBody_out.msg";
            outlookMsg.Save(outputPath);

            Console.WriteLine($"Message:  {outlookMsg.Subject}");
            Console.WriteLine($"RTF body: {outlookMsg.BodyRtf.Length:N0} characters");
            Console.WriteLine($"Saved to {outputPath}");
        }
    }
}

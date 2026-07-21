// Demonstrates how to convert a MIME email message to a MapiMessage and save it as MSG.

using System;
using Aspose.Email.Mapi;

namespace Aspose.Email.Examples.MAPI
{
    internal static class ConvertMimeMessagesFromMsgToEml
    {
        public static void Run()
        {
            var msg = MailMessage.Load(Data.Mapi/"Message.eml");

            // Unicode keeps non-ASCII text intact; the ANSI format would not.
            var mapi = MapiMessage.FromMailMessage(msg, new MapiConversionOptions(OutlookMessageFormat.Unicode));

            var outputPath = Data.Out/"ConvertMIMEMessagesFromMSGToEML_out.msg";
            mapi.Save(outputPath);

            Console.WriteLine($"Converted: {mapi.Subject}");
            Console.WriteLine($"Saved to {outputPath}");
        }
    }
}

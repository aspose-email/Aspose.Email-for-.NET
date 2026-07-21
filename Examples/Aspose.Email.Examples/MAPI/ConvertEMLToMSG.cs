// Demonstrates how to load an EML file and save it as an Outlook MSG file.

using System;

namespace Aspose.Email.Examples.MAPI
{
    internal static class ConvertEmlToMsg
    {
        public static void Run()
        {
            var message = MailMessage.Load(Data.Mapi/"Message.eml", new EmlLoadOptions());

            var outputPath = Data.Out/"ConvertEMLToMSG_out.msg";
            message.Save(outputPath, SaveOptions.DefaultMsgUnicode);

            Console.WriteLine($"Converted: {message.Subject}");
            Console.WriteLine($"Saved to {outputPath}");
        }
    }
}

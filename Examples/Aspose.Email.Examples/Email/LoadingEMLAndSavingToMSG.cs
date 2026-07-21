// Demonstrates how to load an EML file and save it as a Unicode MSG file.

using System;

namespace Aspose.Email.Examples.Email
{
    internal static class LoadingEmlAndSavingToMsg
    {
        public static void Run()
        {
            var eml = MailMessage.Load(Data.Email/"Message.eml");

            // DefaultMsgUnicode keeps non-ASCII text intact; DefaultMsg would write the
            // older ANSI MSG format instead.
            var outputPath = Data.Out/"AnEmail_out.msg";
            eml.Save(outputPath, SaveOptions.DefaultMsgUnicode);

            Console.WriteLine($"Converted: {eml.Subject}");
            Console.WriteLine($"Saved to {outputPath}");
        }
    }
}

// Demonstrates how to save a MailMessage as an Outlook Template (.OFT) file.

using System;

namespace Aspose.Email.Examples.Email
{
    internal static class SaveMessageAsOft
    {
        public static void Run()
        {
            using (var eml = new MailMessage("test@from.to", "test@to.to", "template subject", "Template body"))
            {
                // An OFT is an MSG saved with the template save type, so start from the
                // MSG options and change the save type.
                var options = SaveOptions.DefaultMsgUnicode;
                options.MailMessageSaveType = MailMessageSaveType.OutlookTemplateFormat;

                var outputPath = Data.Out/"EmlAsOft_out.oft";
                eml.Save(outputPath, options);

                Console.WriteLine($"Template: {eml.Subject}");
                Console.WriteLine($"Saved to {outputPath}");
            }
        }
    }
}

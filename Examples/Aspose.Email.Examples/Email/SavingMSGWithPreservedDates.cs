// Demonstrates how to load an EML file and save it as MSG while preserving
// the original send and receive dates.

using System;

namespace Aspose.Email.Examples.Email
{
    internal static class SavingMsgWithPreservedDates
    {
        public static void Run()
        {
            var eml = MailMessage.Load(Data.Email/"Message.eml");

            // Without PreserveOriginalDates the saved MSG would be stamped with the
            // current time instead of the dates the message actually carries.
            var msgSaveOptions = new MsgSaveOptions(MailMessageSaveType.OutlookMessageFormatUnicode)
            {
                PreserveOriginalDates = true
            };

            var outputPath = Data.Out/"outTest_out.msg";
            eml.Save(outputPath, msgSaveOptions);

            Console.WriteLine($"Original date: {eml.Date}");
            Console.WriteLine($"Saved date:    {MailMessage.Load(outputPath).Date}");
            Console.WriteLine($"Saved to {outputPath}");
        }
    }
}

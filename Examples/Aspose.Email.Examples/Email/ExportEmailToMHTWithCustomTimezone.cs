// Demonstrates how to export an EML message to MHTML format with the message date
// rendered using the local system time zone offset.

using System;

namespace Aspose.Email.Examples.Email
{
    internal static class ExportEmailToMhtWithCustomTimezone
    {
        public static void Run()
        {
            var eml = MailMessage.Load(Data.Email/"Message.eml");

            // TimeZoneOffset decides how the Sent date is rendered in the output; without
            // it the date would be written as it is stored in the message.
            eml.TimeZoneOffset = TimeZoneInfo.Local.GetUtcOffset(DateTime.Now);

            var mhtSaveOptions = new MhtSaveOptions { MhtFormatOptions = MhtFormatOptions.WriteHeader };

            var outputPath = Data.Out/"ExportEmailToMHTWithCustomTimezone_out.mhtml";
            eml.Save(outputPath, mhtSaveOptions);

            Console.WriteLine($"Message date:     {eml.Date}");
            Console.WriteLine($"Time zone offset: {eml.TimeZoneOffset}");
            Console.WriteLine($"Saved to {outputPath}");
        }
    }
}

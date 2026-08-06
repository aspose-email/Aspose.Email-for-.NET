// Demonstrates how to change the MHT header markup once for the whole application.
// GlobalFormattingOptions seeds every MhtSaveOptions created afterwards, which saves
// repeating the same templates at each call site.

using System;

namespace Aspose.Email.Examples.Email
{
    internal static class UseGlobalFormattingOptions
    {
        public static void Run()
        {
            // Remember the defaults, so this example does not leak its settings into
            // whatever runs next in the same process.
            var originalPageHeader = GlobalFormattingOptions.PageHeaderFormat;
            var originalHeader = GlobalFormattingOptions.HeaderFormat;
            var originalBefore = GlobalFormattingOptions.BeforeHeadersFormat;
            var originalAfter = GlobalFormattingOptions.AfterHeadersFormat;

            try
            {
                GlobalFormattingOptions.BeforeHeadersFormat = "<div class='aspose-headers'>";
                GlobalFormattingOptions.AfterHeadersFormat = "</div><hr/>";

                // Both instances pick the global settings up at construction time.
                var saveOptions1 = new MhtSaveOptions { MhtFormatOptions = MhtFormatOptions.WriteHeader };
                var saveOptions2 = new MhtSaveOptions { MhtFormatOptions = MhtFormatOptions.WriteHeader };

                Console.WriteLine($"saveOptions1 before-headers: '{saveOptions1.BeforeHeadersFormat}'");
                Console.WriteLine($"saveOptions2 before-headers: '{saveOptions2.BeforeHeadersFormat}'");

                var eml = MailMessage.Load(Data.Email/"Message.eml");
                var outputPath = Data.Out/"UseGlobalFormattingOptions_out.mhtml";
                eml.Save(outputPath, saveOptions1);

                Console.WriteLine($"\nSaved to {outputPath}");
            }
            finally
            {
                GlobalFormattingOptions.PageHeaderFormat = originalPageHeader;
                GlobalFormattingOptions.HeaderFormat = originalHeader;
                GlobalFormattingOptions.BeforeHeadersFormat = originalBefore;
                GlobalFormattingOptions.AfterHeadersFormat = originalAfter;
            }
        }
    }
}

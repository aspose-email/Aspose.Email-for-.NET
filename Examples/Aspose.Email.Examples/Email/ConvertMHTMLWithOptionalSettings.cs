// Demonstrates how to convert an EML message to MHTML format using custom save options,
// including header formatting and skipping inline images.

using System;
using System.IO;

namespace Aspose.Email.Examples.Email
{
    internal static class ConvertMhtmlWithOptionalSettings
    {
        public static void Run()
        {
            var eml = MailMessage.Load(Data.Email/"To be happy.eml");

            // Write headers in the original order without extra print headers,
            // and render addresses as Outlook does.
            var mhtSaveOptions = new MhtSaveOptions
            {
                MhtFormatOptions = MhtFormatOptions.WriteHeader
                                 | MhtFormatOptions.HideExtraPrintHeader
                                 | MhtFormatOptions.DisplayAsOutlook,
                CheckBodyContentEncoding = true
            };

            var withImages = Data.Out/"outMessage_out.mht";
            eml.Save(withImages, mhtSaveOptions);

            // Save again without embedding inline images
            mhtSaveOptions.SkipInlineImages = true;
            var withoutImages = Data.Out/"EmlToMhtmlWithoutInlineImages_out.mht";
            eml.Save(withoutImages, mhtSaveOptions);

            Console.WriteLine($"With inline images:    {new FileInfo(withImages).Length,8:N0} bytes  {withImages}");
            Console.WriteLine($"Without inline images: {new FileInfo(withoutImages).Length,8:N0} bytes  {withoutImages}");
        }
    }
}

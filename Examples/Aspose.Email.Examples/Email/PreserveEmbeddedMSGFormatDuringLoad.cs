// Demonstrates how to preserve the embedded MSG format when loading a TNEF EML file,
// then verify the detected format of the embedded message attachment.

using System;
using Aspose.Email.Tools;

namespace Aspose.Email.Examples.Email
{
    internal static class PreserveEmbeddedMsgFormatDuringLoad
    {
        public static void Run()
        {
            var eml = MailMessage.Load(Data.Email/"tnefWithMsgInside.eml",
                new EmlLoadOptions { PreserveEmbeddedMessageFormat = true });

            var fileFormat = FileFormatUtil
                .DetectFileFormat(eml.Attachments[0].ContentStream)
                .FileFormatType;

            Console.WriteLine($"Embedded message file format: {fileFormat}");
        }
    }
}

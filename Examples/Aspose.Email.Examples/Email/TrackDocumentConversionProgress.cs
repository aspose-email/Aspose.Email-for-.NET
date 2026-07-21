// Demonstrates how to track EML conversion progress using a custom progress handler
// that reports MIME structure creation and part-saving events.

using System;
using System.IO;

namespace Aspose.Email.Examples.Email
{
    internal static class TrackDocumentConversionProgress
    {
        public static void Run()
        {
            var eml = MailMessage.Load(Data.Email/"test.eml");

            using (var ms = new MemoryStream())
            {
                var opt = new EmlSaveOptions(MailMessageSaveType.EmlFormat)
                {
                    CustomProgressHandler = ShowEmlConversionProgress
                };

                eml.Save(ms, opt);
            }
        }

        private static void ShowEmlConversionProgress(ProgressEventHandlerInfo info)
        {
            switch (info.EventType)
            {
                case ProgressEventType.MimeStructureCreated:
                    Console.WriteLine($"MimeStructureCreated - Total: {info.TotalMimePartCount}, Saved: {info.SavedMimePartCount}");
                    break;
                case ProgressEventType.MimePartSaved:
                    Console.WriteLine($"MimePartSaved - Total: {info.TotalMimePartCount}, Saved: {info.SavedMimePartCount}");
                    break;
                case ProgressEventType.SavedToStream:
                    Console.WriteLine($"SavedToStream - Total: {info.TotalMimePartCount}, Saved: {info.SavedMimePartCount}");
                    break;
            }
        }
    }
}

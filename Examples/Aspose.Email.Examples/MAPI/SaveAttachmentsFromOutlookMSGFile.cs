// Demonstrates how to save every attachment of an Outlook MSG file to disk.

using System;
using Aspose.Email.Mapi;

namespace Aspose.Email.Examples.MAPI
{
    internal static class SaveAttachmentsFromOutlookMSGFile
    {
        public static void Run()
        {
            var outputDir = Data.OutSub("Attachments");
            var message = MapiMessage.Load(Data.Mapi/"outputAttachments.msg");

            foreach (var attachment in message.Attachments)
            {
                attachment.Save(outputDir/attachment.FileName);
                Console.WriteLine($"Saved: {attachment.FileName}");
            }

            Console.WriteLine($"\nSaved {message.Attachments.Count} attachment(s) to {outputDir}");
        }
    }
}

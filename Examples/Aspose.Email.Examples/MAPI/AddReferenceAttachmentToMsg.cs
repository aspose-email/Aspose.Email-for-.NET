// Demonstrates how to attach a cloud link instead of a file. A reference attachment
// carries no content, only the URL and the permissions the recipient gets.

using System;
using Aspose.Email.Mapi;

namespace Aspose.Email.Examples.MAPI
{
    internal static class AddReferenceAttachmentToMsg
    {
        public static void Run()
        {
            var msg = new MapiMessage("from@domain.com", "to@domain.com", "Outlook message file",
                "This message is created by Aspose.Email", OutlookMessageFormat.Unicode);

            var options = new ReferenceAttachmentOptions(
                "https://drive.google.com/file/d/1HJ-M3F2qq1oRrTZ2GZhUdErJNy2CT3DF/",
                "https://drive.google.com/drive/my-drive",
                "GoogleDrive")
            {
                PermissionType = AttachmentPermissionType.AnyoneCanEdit,
                IsFolder = false
            };

            msg.Attachments.Add("Document.pdf", options);

            // Add an ordinary attachment too, so the IsReference check has both cases.
            msg.Attachments.Add("1.txt", System.IO.File.ReadAllBytes(Data.Mapi/"1.txt"));

            foreach (var attachment in msg.Attachments)
            {
                Console.WriteLine($"{attachment.LongFileName}: " +
                                  (attachment.IsReference ? "reference attachment" : "regular attachment"));
            }

            var outputPath = Data.Out/"AddReferenceAttachmentToMsg_out.msg";
            msg.Save(outputPath);
            Console.WriteLine($"\nSaved to {outputPath}");
        }
    }
}

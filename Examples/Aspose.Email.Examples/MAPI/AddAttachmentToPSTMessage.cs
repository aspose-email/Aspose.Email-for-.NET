// Demonstrates how to attach a file to a message that already sits inside a PST,
// without extracting, rebuilding and re-adding the whole message.

using System;
using System.IO;
using Aspose.Email.Storage.Pst;

namespace Aspose.Email.Examples.MAPI
{
    internal static class AddAttachmentToPSTMessage
    {
        public static void Run()
        {
            // Work on a copy: the attachment is written into the storage.
            var pstPath = Data.Out/"AddAttachmentToPSTMessage_out.pst";
            File.Copy(Data.Mapi/"Sub.pst", pstPath, true);

            using (var pst = PersonalStorage.FromFile(pstPath))
            {
                var inbox = pst.RootFolder.GetSubFolder("Inbox");
                var messageInfo = inbox.GetContents()[0];

                Console.WriteLine($"Message: {messageInfo.Subject}");
                Console.WriteLine($"Attachments before: {pst.ExtractMessage(messageInfo).Attachments.Count}");

                pst.AddAttachmentToMessage(messageInfo, Data.Mapi/"1.txt");

                // Adding an attachment rewrites the message, so look it up again to
                // read the updated item rather than reusing the stale MessageInfo.
                var updated = pst.RootFolder.GetSubFolder("Inbox").GetContents()[0];
                var message = pst.ExtractMessage(updated);

                Console.WriteLine($"Attachments after:  {message.Attachments.Count}");
                foreach (var attachment in message.Attachments)
                    Console.WriteLine($"  {attachment.LongFileName}");
            }

            Console.WriteLine($"\nSaved to {pstPath}");
        }
    }
}

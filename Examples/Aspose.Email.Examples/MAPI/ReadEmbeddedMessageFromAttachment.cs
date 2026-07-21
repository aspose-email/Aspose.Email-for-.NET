// Demonstrates how to detect and extract an embedded MapiMessage from an MSG attachment.

using System;
using Aspose.Email.Mapi;

namespace Aspose.Email.Examples.MAPI
{
    internal static class ReadEmbeddedMessageFromAttachment
    {
        public static void Run()
        {
            var message = MapiMessage.Load(Data.Mapi/"WithEmbeddedMsg.msg");
            if (message.Attachments[0].ObjectData.IsOutlookMessage)
            {
                var embedded = message.Attachments[0].ObjectData.ToMapiMessage();
                Console.WriteLine("Embedded subject: " + embedded.Subject);
            }
        }
    }
}

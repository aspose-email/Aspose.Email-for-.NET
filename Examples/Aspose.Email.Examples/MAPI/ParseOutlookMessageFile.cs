// Demonstrates how to parse an Outlook EML file and display its subject, sender, body, and attachments.

using System;
using Aspose.Email.Mapi;

namespace Aspose.Email.Examples.MAPI
{
    internal static class ParseOutlookMessageFile
    {
        public static void Run()
        {
            MapiMessage msg = MapiMessage.FromMailMessage(Data.Mapi/"Message.eml");

            Console.WriteLine("Subject:" + msg.Subject);
            Console.WriteLine("From:" + msg.SenderName);
            Console.WriteLine("Body:" + msg.Body);
            Console.WriteLine("Attachment Count:" + msg.Attachments.Count);

            foreach (MapiAttachment attachment in msg.Attachments)
            {
                Console.WriteLine("Attachment:" + attachment.FileName);
                attachment.Save(Data.Out/attachment.FileName);
            }
        }
    }
}

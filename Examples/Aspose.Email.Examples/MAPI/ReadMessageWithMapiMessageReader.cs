// Demonstrates MapiMessageReader, a lighter way into an MSG file than MapiMessage.Load.
//
// ReadAttachments pulls just the attachment table, so a job that only needs the files
// out of a message never pays for parsing the body, recipients and named properties.
// ReadMessage is there when the whole message is wanted after all.

using System;
using Aspose.Email.Mapi;

namespace Aspose.Email.Examples.MAPI
{
    internal static class ReadMessageWithMapiMessageReader
    {
        public static void Run()
        {
            var outputDir = Data.OutSub("ReaderAttachments");

            foreach (var fileName in new[] { "MsgWithAtt.msg", "WithEmbeddedMsg.msg", "message.msg" })
            {
                Console.WriteLine($"--- {fileName} ---");

                // Attachments only.
                using (var reader = new MapiMessageReader(Data.Mapi/fileName))
                {
                    var attachments = reader.ReadAttachments();
                    Console.WriteLine($"  attachments: {attachments.Count}");

                    foreach (var attachment in attachments)
                    {
                        // An OLE attachment such as an embedded message carries no
                        // BinaryData - its content lives in ObjectData instead.
                        var size = attachment.BinaryData == null ? 0 : attachment.BinaryData.Length;
                        Console.WriteLine($"    {attachment.LongFileName} ({size:N0} bytes)");

                        if (size > 0)
                            attachment.Save(outputDir/attachment.LongFileName);
                    }
                }

                // The same reader can hand back the whole message when it is needed.
                using (var reader = new MapiMessageReader(Data.Mapi/fileName))
                {
                    var message = reader.ReadMessage();
                    Console.WriteLine($"  subject: {message.Subject}");
                    Console.WriteLine($"  sender:  {message.SenderName}");
                }
            }

            Console.WriteLine($"\nSaved attachments to {outputDir}");
        }
    }
}

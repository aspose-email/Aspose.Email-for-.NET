// Demonstrates how to tell an OLE attachment from an ordinary one and read what it
// wraps.
//
// A plain file attachment keeps its bytes in BinaryData and leaves ObjectData null.
// An OLE attachment - an embedded message, a linked Word or Excel document - carries a
// MapiObjectProperty instead, which names the producing application and, for an
// embedded Outlook item, converts straight to a MapiMessage.

using System;
using Aspose.Email.Mapi;

namespace Aspose.Email.Examples.MAPI
{
    internal static class InspectOleAttachmentProperties
    {
        public static void Run()
        {
            foreach (var fileName in new[] { "WithEmbeddedMsg.msg", "messageWithEmbeddedEML.msg", "MsgWithAtt.msg" })
            {
                var msg = MapiMessage.Load(Data.Mapi/fileName);
                Console.WriteLine($"--- {fileName}: {msg.Attachments.Count} attachment(s) ---");

                foreach (var attachment in msg.Attachments)
                {
                    var objectData = attachment.ObjectData;

                    if (objectData == null)
                    {
                        var size = attachment.BinaryData == null ? 0 : attachment.BinaryData.Length;
                        Console.WriteLine($"  {attachment.LongFileName}: plain attachment, {size:N0} bytes");
                        continue;
                    }

                    Console.WriteLine($"  {attachment.LongFileName}: OLE object");
                    Console.WriteLine($"    document name:  {Show(objectData.DocumentName)}");
                    Console.WriteLine($"    producer:       {objectData.OleDocumentFormat}");
                    Console.WriteLine($"    Outlook item:   {objectData.IsOutlookMessage}");
                    Console.WriteLine($"    properties:     {objectData.Properties.Count}");

                    // Comparing against the well-known formats avoids parsing the GUID.
                    if (Equals(objectData.OleDocumentFormat, OleDocumentFormat.MicrosoftOfficeWordDocument))
                        Console.WriteLine("    -> a Word document");

                    // Only an embedded Outlook item converts to a message.
                    if (objectData.IsOutlookMessage)
                    {
                        var embedded = objectData.ToMapiMessage();
                        Console.WriteLine($"    embedded subject: {embedded.Subject}");
                        Console.WriteLine($"    embedded sender:  {embedded.SenderName}");

                        var outputPath = Data.Out/"InspectOleAttachmentProperties_embedded.msg";
                        embedded.Save(outputPath);
                        Console.WriteLine($"    saved to {outputPath}");
                    }
                }
            }
        }

        private static string Show(string value)
        {
            return string.IsNullOrEmpty(value) ? "(not set)" : value;
        }
    }
}

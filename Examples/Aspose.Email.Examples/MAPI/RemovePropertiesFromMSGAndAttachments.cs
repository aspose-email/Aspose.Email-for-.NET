// Demonstrates how to remove a MAPI property from an attachment, and how to check
// whether the removal survived a save and reload.

using System;
using Aspose.Email.Mapi;

namespace Aspose.Email.Examples.MAPI
{
    internal static class RemovePropertiesFromMSGAndAttachments
    {
        // PR_ATTACH_LONG_PATHNAME_W - the attachment's original full path.
        private const long PropertyToRemove = 923467779;

        public static void Run()
        {
            var mapi = new MapiMessage("from@doamin.com", "to@domain.com", "subject", "body");
            mapi.SetBodyContent("<html><body><h1>This is the body content</h1></body></html>", BodyContentType.Html);

            var attachment = MapiMessage.Load(Data.Mapi/"message.msg");
            mapi.Attachments.Add("Outlook2 Test subject.msg", attachment);

            var lastAttachment = mapi.Attachments[mapi.Attachments.Count - 1];
            Console.WriteLine($"Properties before removal: {lastAttachment.Properties.Count}");

            lastAttachment.RemoveProperty(PropertyToRemove);
            Console.WriteLine($"Properties after removal:  {lastAttachment.Properties.Count}");

            var outputPath = Data.Out/"EMAIL_589265.msg";
            mapi.Save(outputPath);

            // The removal applies to the in-memory message. Saving regenerates the
            // attachment's standard properties, so the count is back up after a reload -
            // remove the property again after loading if it must stay out.
            var reloaded = MapiMessage.Load(outputPath);
            Console.WriteLine($"Properties after reload:   {reloaded.Attachments[reloaded.Attachments.Count - 1].Properties.Count}");
            Console.WriteLine($"Saved to {outputPath}");
        }
    }
}

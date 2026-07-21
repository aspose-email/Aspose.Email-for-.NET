// Demonstrates how to distinguish between inline and regular attachments in an MSG message
// by inspecting MAPI properties.

using System;
using Aspose.Email.Mapi;

namespace Aspose.Email.Examples.Email
{
    internal static class IdentifyInlineAndRegularAttachments
    {
        public static void Run()
        {
            var eml = MapiMessage.Load(Data.Email/"Important decision.msg");

            Console.WriteLine($"Total attachments: {eml.Attachments.Count}");

            foreach (var attachment in eml.Attachments)
            {
                Console.WriteLine(IsInlineAttachment(attachment, eml)
                    ? $"{attachment.LongFileName} is inline attachment"
                    : $"{attachment.LongFileName} is regular attachment");
            }
        }

        private static bool IsInlineAttachment(MapiAttachment att, MapiMessage msg)
        {
            switch (msg.BodyType)
            {
                case BodyContentType.PlainText:
                    return false;

                case BodyContentType.Html:
                    if (att.Properties.ContainsKey(0x37140003))
                    {
                        var attachFlagsValue = att.GetPropertyLong(0x37140003);
                        if (attachFlagsValue != null && (attachFlagsValue & 0x00000004) == 0x00000004)
                        {
                            // Check PidTagAttachContentId
                            if (att.Properties.ContainsKey(MapiPropertyTag.PR_ATTACH_CONTENT_ID) ||
                                att.Properties.ContainsKey(MapiPropertyTag.PR_ATTACH_CONTENT_ID_W))
                            {
                                var contentId = att.Properties.ContainsKey(MapiPropertyTag.PR_ATTACH_CONTENT_ID)
                                    ? att.Properties[MapiPropertyTag.PR_ATTACH_CONTENT_ID].GetString()
                                    : att.Properties[MapiPropertyTag.PR_ATTACH_CONTENT_ID_W].GetString();
                                if (msg.BodyHtml.Contains(contentId))
                                    return true;
                            }
                            // Check PidTagAttachContentLocation
                            if (att.Properties.ContainsKey(0x3713001E) || att.Properties.ContainsKey(0x3713001F))
                                return true;
                        }
                        else if (att.Properties.ContainsKey(0x3716001F) && att.GetPropertyString(0x3716001F) == "inline"
                              || att.Properties.ContainsKey(0x3716001E) && att.GetPropertyString(0x3716001E) == "inline")
                        {
                            return true;
                        }
                    }
                    else if (att.Properties.ContainsKey(0x3716001F) && att.GetPropertyString(0x3716001F) == "inline"
                          || att.Properties.ContainsKey(0x3716001E) && att.GetPropertyString(0x3716001E) == "inline")
                    {
                        return true;
                    }
                    return false;

                case BodyContentType.Rtf:
                    // For RTF bodies, all OLE attachments (PidTagAttachMethod = 0x00000006) are inline
                    if (att.Properties.ContainsKey(MapiPropertyTag.PR_ATTACH_METHOD))
                        return att.GetPropertyLong(MapiPropertyTag.PR_ATTACH_METHOD) == 0x00000006;
                    return false;

                default:
                    throw new ArgumentOutOfRangeException();
            }
        }
    }
}

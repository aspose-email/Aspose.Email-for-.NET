using System;
using System.IO;
using System.Text;
using Aspose.Email.AntiSpam;
using Aspose.Email.Mime;
using Aspose.Email.Mapi;


/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/

namespace Aspose.Email.Examples.CSharp.Email
{
    class IdentifyInlineAndRegularAttachments
    {
        // ExStart:IdentifyInlineAndRegularAttachments
        public static void Run()
        {                       
            string dataDir = RunExamples.GetDataDir_Email();
            var message = MapiMessage.FromFile(dataDir + "RemoveAttachments.msg");
            var attachments = message.Attachments;
            var count = attachments.Count;

            Console.WriteLine("Total attachments " + count);
            for (int i = 0; i < attachments.Count; i++)
            {
                var attachment = attachments[i];
                if (IsInlineAttachment(attachment, message))
                {
                    Console.WriteLine(attachment.LongFileName + " is inline attachment");
                }
                else
                {
                    Console.WriteLine(attachment.LongFileName + " is regular attachment");
                }
            }

        }

        public static bool IsInlineAttachment(MapiAttachment att, MapiMessage msg)
        {
            switch (msg.BodyType)
            {
                case BodyContentType.PlainText:
                    // ignore indications for plain text messages
                    return false;

                case BodyContentType.Html:

                    // check the PidTagAttachFlags property
                    if (att.Properties.ContainsKey(0x37140003))
                    {
                        long? attachFlagsValue = att.GetPropertyLong(0x37140003);
                        if (attachFlagsValue != null && (attachFlagsValue & 0x00000004) == 0x00000004)
                        {
                            // check PidTagAttachContentId property
                            if (att.Properties.ContainsKey(MapiPropertyTag.PR_ATTACH_CONTENT_ID) ||
                                att.Properties.ContainsKey(MapiPropertyTag.PR_ATTACH_CONTENT_ID_W))
                            {
                                string contentId = att.Properties.ContainsKey(MapiPropertyTag.PR_ATTACH_CONTENT_ID)
                                    ? att.Properties[MapiPropertyTag.PR_ATTACH_CONTENT_ID].GetString()
                                    : att.Properties[MapiPropertyTag.PR_ATTACH_CONTENT_ID_W].GetString();
                                if (msg.BodyHtml.Contains(contentId))
                                {
                                    return true;
                                }
                            }
                            // check PidTagAttachContentLocation property
                            if (att.Properties.ContainsKey(0x3713001E) ||
                                att.Properties.ContainsKey(0x3713001F))
                            {
                                return true;
                            }
                        }
                        else if ((att.Properties.ContainsKey(0x3716001F) && att.GetPropertyString(0x3716001F) == "inline")
                            || (att.Properties.ContainsKey(0x3716001E) && att.GetPropertyString(0x3716001E) == "inline"))
                        {
                            return true;
                        }
                    }
                    else if ((att.Properties.ContainsKey(0x3716001F) && att.GetPropertyString(0x3716001F) == "inline")
                            || (att.Properties.ContainsKey(0x3716001E) && att.GetPropertyString(0x3716001E) == "inline"))
                    {
                        return true;
                    }
                    return false;

                case BodyContentType.Rtf:

                    // If the body is RTF, then all OLE attachments are inline attachments.
                    // OLE attachments have 0x00000006 for the value of the PidTagAttachMethod property
                    if (att.Properties.ContainsKey(MapiPropertyTag.PR_ATTACH_METHOD))
                    {
                        return att.GetPropertyLong(MapiPropertyTag.PR_ATTACH_METHOD) == 0x00000006;
                    }
                    return false;
                default:
                    throw new ArgumentOutOfRangeException();
            }
        }
        // ExEnd:IdentifyInlineAndRegularAttachments
    }
}

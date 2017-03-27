using Aspose.Email.Mime;
using Aspose.Email.Mapi;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/

namespace Aspose.Email.Examples.CSharp.Email.Outlook
{
    class GetNestedMailMessageAttachments
    {
        public static void Run()
        {
            string dataDir = RunExamples.GetDataDir_Outlook();
            MapiMessage message = MapiMessage.FromFile(dataDir + "messageWithEmbeddedEML.msg");
            MapiAttachment attachment = message.Attachments[0];

            // ExStart:GetNestedMailMessageAttachments
            // Create a MapiMessage object from the individual attachment
            MapiMessage getAttachment = MapiMessage.FromProperties(attachment.ObjectData.Properties);

            // Create object of type MailMessageImterpretor from the above message and Save the embedded message to file at disk
            MailMessage mailMessage = getAttachment.ToMailMessage(new MailConversionOptions());
            mailMessage.Save(dataDir + @"NestedMailMessageAttachments_out.eml", SaveOptions.DefaultEml);
            // ExEnd:GetNestedMailMessageAttachments
        }
    }
}

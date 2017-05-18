using Aspose.Email.Mapi;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

/* This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET 
   API reference when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq 
   for more information. If you do not wish to use NuGet, you can manually download 
   Aspose.Email for .NET API from http://www.aspose.com/downloads, 
   install it and then add its reference to this project. For any issues, questions or suggestions 
   please feel free to contact us using http://www.aspose.com/community/forums/default.aspx            
*/

namespace Aspose.Email.Examples.CSharp.Email.Outlook
{
    class RenderingContactInformationToMhtml
    {
        public static void Run()
        {
            // ExStart:RenderingContactInformationToMhtml
            string dataDir = RunExamples.GetDataDir_Outlook();

            //Load VCF Contact and convert to MailMessage for rendering to MHTML
            MapiContact contact = MapiContact.FromVCard(dataDir + "Contact.vcf");

            MemoryStream ms = new MemoryStream();
            contact.Save(ms, ContactSaveFormat.Msg);
            ms.Position = 0;
            MapiMessage msg = MapiMessage.FromStream(ms);
            MailConversionOptions op = new MailConversionOptions();
            MailMessage eml = msg.ToMailMessage(op);

            //Prepare the MHT format options
            MhtSaveOptions mhtSaveOptions = new MhtSaveOptions();
            mhtSaveOptions.CheckBodyContentEncoding = true;
            mhtSaveOptions.PreserveOriginalBoundaries = true;
            MhtFormatOptions formatOp = MhtFormatOptions.WriteHeader | MhtFormatOptions.RenderVCardInfo;
            mhtSaveOptions.RenderedContactFields = ContactFieldsSet.NameInfo | ContactFieldsSet.PersonalInfo | ContactFieldsSet.Telephones | ContactFieldsSet.Events;
            mhtSaveOptions.MhtFormatOptions = formatOp;
            eml.Save(dataDir + "ContactMhtml_out.mhtml", mhtSaveOptions);
            // ExEnd:RenderingContactInformationToMhtml
        }
    }
}

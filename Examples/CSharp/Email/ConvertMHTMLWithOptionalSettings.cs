using Aspose.Email.Mime;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/

namespace Aspose.Email.Examples.CSharp.Email
{
    class ConvertMHTMLWithOptionalSettings
    {
        public static void Run()
        {
            string dataDir = RunExamples.GetDataDir_Email();

            // ExStart:ConvertMHTMLWithOptionalSettings
            MailMessage eml = MailMessage.Load(Path.Combine(dataDir, "Message.eml"));

            // Save as mht with header
            MhtSaveOptions mhtSaveOptions = new MhtSaveOptions
            {
                //Specify formatting options required
                //Here we are specifying to write header informations to outpu without writing extra print header
                //and the output headers should display as the original headers in message
                MhtFormatOptions = MhtFormatOptions.WriteHeader | MhtFormatOptions.HideExtraPrintHeader | MhtFormatOptions.DisplayAsOutlook,
                // Check the body encoding for validity. 
                CheckBodyContentEncoding = true
            };
            eml.Save(Path.Combine(dataDir, "outMessage_out.mht"), mhtSaveOptions);
            // ExEnd:ConvertMHTMLWithOptionalSettings

            //ExStart: ConvertToMHTMLWithoutInlineImages
            mhtSaveOptions.SkipInlineImages = true;

            eml.Save(Path.Combine(dataDir, "EmlToMhtmlWithoutInlineImages_out.mht"), mhtSaveOptions);
            //ExEnd: ConvertToMHTMLWithoutInlineImages
        }
    }
}

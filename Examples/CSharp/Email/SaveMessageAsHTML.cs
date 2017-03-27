using Aspose.Email.Mime;
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

namespace Aspose.Email.Examples.CSharp.Email
{
    class SaveMessageAsHTML
    {
        public static void Run()
        {
            string dataDir = RunExamples.GetDataDir_Email();

            // ExStart:SaveMessageAsHTML
            MailMessage message = MailMessage.Load(dataDir + "Message.eml");
            message.Save(dataDir + "SaveAsHTML_out.html", SaveOptions.DefaultHtml);

            // OR

            MailMessage eml = MailMessage.Load(dataDir + "Message.eml");
            HtmlSaveOptions options = SaveOptions.DefaultHtml;
            options.EmbedResources = false;
            options.HtmlFormatOptions = HtmlFormatOptions.WriteHeader | HtmlFormatOptions.WriteCompleteEmailAddress; //save the message headers to output HTML using the formatting options
            eml.Save(dataDir + "SaveAsHTML1_out.html", options);
            // ExEnd:SaveMessageAsHTML
        }
    }
}

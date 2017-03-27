using System.Text;
using Aspose.Email.Mime;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/

namespace Aspose.Email.Examples.CSharp.Email
{
    class LoadMessageWithLoadOptions
    {
        public static void Run()
        {
            // ExStart:LoadMessageWithLoadOptions
            // The path to the File directory.
            string dataDir = RunExamples.GetDataDir_Email();

            // Load Eml, html, mhtml, msg and dat file 
            MailMessage mailMessage = MailMessage.Load(dataDir + "Message.eml", new EmlLoadOptions());
            MailMessage.Load(dataDir + "description.html", new HtmlLoadOptions());
            MailMessage.Load(dataDir + "Message.mhtml", new MhtmlLoadOptions());
            MailMessage.Load(dataDir + "Message.msg", new MsgLoadOptions());

            // loading with custom options
            EmlLoadOptions emlLoadOptions = new EmlLoadOptions
            {
                PrefferedTextEncoding = Encoding.UTF8,
                PreserveTnefAttachments = true
            };

            MailMessage.Load(dataDir + "description.html", emlLoadOptions);
            HtmlLoadOptions htmlLoadOptions = new HtmlLoadOptions
            {
                PrefferedTextEncoding = Encoding.UTF8,
                ShouldAddPlainTextView = true,
                PathToResources = dataDir
            };
            MailMessage.Load(dataDir + "description.html", emlLoadOptions);
            // ExEnd:LoadMessageWithLoadOptions
        }
    }
}

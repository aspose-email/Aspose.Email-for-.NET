using Aspose.Email.Mime;
using System;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from https://www.nuget.org/packages/Aspose.Email/, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using https://forum.aspose.com/c/email
*/

namespace Aspose.Email.Examples.CSharp.Email
{
    class GetHTMLBodyAsPlainText
    {
        public static void Run()
        {
            // ExStart:1
            // The path to the File directory.
            string dataDir = RunExamples.GetDataDir_Email();

            MailMessage mail = MailMessage.Load(dataDir + "HtmlWithUrlSample.eml");

            string body_with_url = mail.GetHtmlBodyText(true);// body will contain URL
            string body_without_url = mail.GetHtmlBodyText(false);// body will not contain URL

            Console.WriteLine("Body with URL: " + body_with_url);
            Console.WriteLine("Body without URL: " + body_without_url);
            // ExEnd:1
        }
    }
}

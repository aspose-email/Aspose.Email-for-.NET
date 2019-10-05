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
    class CustomHyperlinkRendering
    {
        // ExStart:1
        public static void Run()
        {
            string dataDir = RunExamples.GetDataDir_Email();
            var fileName = dataDir + "LinksSample.eml";
            MailMessage msg = MailMessage.Load(fileName);
            Console.WriteLine(msg.GetHtmlBodyText(RenderHyperlinkWithHref));

            Console.WriteLine(msg.GetHtmlBodyText(RenderHyperlinkWithoutHref));
        }

        private static string RenderHyperlinkWithHref(string source)
        {
            int start = source.IndexOf("href=\"") + "href=\"".Length;
            int end = source.IndexOf("\"", start + "href=\"".Length);
            string href = source.Substring(start, end - start);
            start = source.IndexOf(">") + 1;
            end = source.IndexOf("<", start);
            string text = source.Substring(start, end - start);
            string link = string.Format("{0}<{1}>", text, href);
            return link;
        }

        private static string RenderHyperlinkWithoutHref(string source)
        {
            int start = source.IndexOf(">") + 1;
            int end = source.IndexOf("<", start);
            string text = source.Substring(start, end - start);
            return text;
        }
        // ExEnd:1
    }
}

using System;
using System.IO;
using System.Security.Cryptography.X509Certificates;
using Aspose.Email.Mail;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/

namespace Aspose.Email.Examples.CSharp.Email
{
    class PrintHeaderUsingMhtFormatOptions
    {
        public static void Run()
        {
            // ExStart:PrintHeaderUsingMhtFormatOptions
            // The path to the File directory.
            string dataDir = RunExamples.GetDataDir_Email();
            const string pageHeader = @"<div><div class='pageHeader'>&quot;Panditharatne, Mithra&quot; &lt;mithra.panditharatne@cibc.com&gt;<hr/></div>";
            MailMessage message = MailMessage.Load(dataDir + "Message.eml");
            MhtMessageFormatter mailFormatter = new MhtMessageFormatter();
            MailMessage copyMessage = message.Clone();
            mailFormatter.Format(copyMessage);
            Console.WriteLine(copyMessage.HtmlBody.Contains(pageHeader) ? "True" : "False");
            MhtFormatOptions options = MhtFormatOptions.HideExtraPrintHeader | MhtFormatOptions.WriteCompleteEmailAddress;
            mailFormatter.Format(message, options);
            if (!message.HtmlBody.Contains(pageHeader))
                Console.WriteLine("True");
            else
                Console.WriteLine("False");
            // ExEnd:PrintHeaderUsingMhtFormatOptions
        }
    }
}

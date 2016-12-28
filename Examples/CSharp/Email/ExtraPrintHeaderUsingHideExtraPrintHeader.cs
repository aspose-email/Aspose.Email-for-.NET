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
    class ExtraPrintHeaderUsingHideExtraPrintHeader
    {
        public static void Run()
        {
            // ExStart:ExtraPrintHeaderUsingHideExtraPrintHeader
            // The path to the File directory.
            string dataDir = RunExamples.GetDataDir_Email();

            string mhtFileName = dataDir + "Message.mhtml";
            MailMessage message = MailMessage.Load(dataDir + "Message.eml");
            string encodedPageHeader = @"<div><div class=3D'page=Header'>&quot;Panditharatne, Mithra&quot; &lt;mithra=2Epanditharatne@cibc==2Ecom&gt;<hr/></div>";

            MhtMessageFormatter mailFormatter = new MhtMessageFormatter();
            MhtFormatOptions options = MhtFormatOptions.WriteCompleteEmailAddress | MhtFormatOptions.WriteHeader;
            mailFormatter.Format(message);

            message.Save(mhtFileName, Aspose.Email.Mail.SaveOptions.DefaultMhtml);

            if (File.ReadAllText(mhtFileName).Contains(encodedPageHeader))
            {
                Console.WriteLine("True");
            }
            else
            {
                Console.WriteLine("False");
            }

            //Assert.True(File.ReadAllText(mhtFileName).Contains(encodedPageHeader));
            options = options | MhtFormatOptions.HideExtraPrintHeader;
            mailFormatter.Format(message);
            message.Save(mhtFileName, Aspose.Email.Mail.SaveOptions.DefaultMhtml);
            if (File.ReadAllText(mhtFileName).Contains(encodedPageHeader))
            {
                Console.WriteLine("True");
            }
            else
            {
                Console.WriteLine("False");
            }
            // ExEnd:ExtraPrintHeaderUsingHideExtraPrintHeader
        }
    }
}

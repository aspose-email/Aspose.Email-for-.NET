using Aspose.Email.Clients.Smtp;
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

namespace Aspose.Email.Examples.CSharp.Email.SMTP
{
    class LoadingEMLFilesFromDisk
    {
        public static void Run()
        {
            string dataDir = RunExamples.GetDataDir_SMTP();

            // ExStart:LoadingEMLFilesFromDisk
            // Load an EML file in MailMessage class
            MailMessage message = MailMessage.Load(dataDir + "test.eml");

            // Send this message using SmtpClient
            SmtpClient client = new SmtpClient("host", "username", "password");
            
            try
            {
                client.Send(message);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }            
            // ExEnd:LoadingEMLFilesFromDisk
        }
    }
}

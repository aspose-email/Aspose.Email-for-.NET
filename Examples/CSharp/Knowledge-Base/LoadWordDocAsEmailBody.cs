using Aspose.Email.Clients;
using Aspose.Email.Clients.Smtp;
using Aspose.Words;
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

namespace Aspose.Email.Examples.CSharp.Email.Knowledge.Base
{
    class LoadWordDocAsEmailBody
    {
        public static void Run()
        {
            try
            {
                // ExStart:LoadWordDocAsEmailBody
                // The path to the File directory
                string dataDir = RunExamples.GetDataDir_KnowledgeBase();

                // Load a Word document from disk and save to stream as MHTML
                Document wordDocument = new Document(dataDir + "Test.doc");
                MemoryStream mhtmlStream = new MemoryStream();
                wordDocument.Save(mhtmlStream, SaveFormat.Mhtml);

                // Load the MHTML in MailMessage
                mhtmlStream.Position = 0;
                MailMessage message = MailMessage.Load(mhtmlStream, new MhtmlLoadOptions());
                message.Subject = "Sending Invoice in Email";
                message.From = "sender@gmail.com";
                message.To = "recipient@gmail.com";

                // Save the message in MSG format to disk
                message.Save(dataDir + "WordDocAsEmailBody_out.msg", SaveOptions.DefaultMsgUnicode);

                // Send in an email
                SmtpClient client = new SmtpClient("smtp.gmail.com", 587, "sender@gmail.com", "pwd");
                client.SecurityOptions = SecurityOptions.SSLExplicit;                
                client.Send(message);
                // ExEnd:LoadWordDocAsEmailBody
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
    }
}

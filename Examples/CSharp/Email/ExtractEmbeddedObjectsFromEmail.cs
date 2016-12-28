using System;
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
    class ExtractEmbeddedObjectsFromEmail
    {
        public static void Run()
        {
            // ExStart:ExtractEmbeddedObjectsFromEmail
            // The path to the File directory.
            string dataDir = RunExamples.GetDataDir_Email();

            // Create an instance of MailMessage and load an email file
            MailMessage mailMsg = MailMessage.Load(dataDir + "Message.msg", new MsgLoadOptions());

            foreach (Attachment attachment in mailMsg.Attachments)
            {
                // To display the the attachment file name
                attachment.Save(dataDir + "MessageEmbedded_out.msg");
                Console.WriteLine(attachment.Name);
            }

            mailMsg.From = "sender@sender.com";
            mailMsg.To.Add("receiver@receiver.com");

            SmtpClient client = new SmtpClient("smtp.server.com")
            {
                Port = 25,
                Username = "UserName",
                Password = "password"
            };

            try
            {
                client.Send(dataDir + "Message.msg");
            }
            catch (Exception ex)
            {
                
               Console.Write(ex.Message);
            }
          
            // ExEnd:ExtractEmbeddedObjectsFromEmail
        }
    }
}
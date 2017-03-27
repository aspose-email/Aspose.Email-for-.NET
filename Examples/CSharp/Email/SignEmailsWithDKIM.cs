using System.IO;
using System.Security.Cryptography;
using Aspose.Email.Mime;
using Aspose.Email.DKIM;
using Aspose.Email.Clients.Smtp;

/* This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET 
   API reference when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq 
   for more information. If you do not wish to use NuGet, you can manually download 
   Aspose.Email for .NET API from http://www.aspose.com/downloads, 
   install it and then add its reference to this project. For any issues, questions or suggestions 
   please feel free to contact us using http://www.aspose.com/community/forums/default.aspx            
*/

namespace Aspose.Email.Examples.CSharp.Email
{
    class SignEmailsWithDKIM
    {
        public static void Run()
        {
            // ExStart:SignEmailsWithDKIM
            string privateKeyFile = Path.Combine(RunExamples.GetDataDir_SMTP().Replace("_Send", string.Empty), RunExamples.GetDataDir_SMTP()+ "key2.pem");

            RSACryptoServiceProvider rsa = PemReader.GetPrivateKey(privateKeyFile);
            DKIMSignatureInfo signInfo = new DKIMSignatureInfo("test", "yandex.ru");
            signInfo.Headers.Add("From");
            signInfo.Headers.Add("Subject");

            MailMessage mailMessage = new MailMessage("useremail@gmail.com", "test@gmail.com");
            mailMessage.Subject = "Signed DKIM message text body";
            mailMessage.Body = "This is a text body signed DKIM message";
            MailMessage signedMsg = mailMessage.DKIMSign(rsa, signInfo);

            try
            {
                SmtpClient client = new SmtpClient("smtp.gmail.com", 587, "your.email@gmail.com", "your.password");
                client.Send(signedMsg);                
            }
            finally
            {}
            // ExEnd:SignEmailsWithDKIM
        }
    }
}

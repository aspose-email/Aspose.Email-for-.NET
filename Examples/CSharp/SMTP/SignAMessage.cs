using System.Security.Cryptography.X509Certificates;
using Aspose.Email.Mime;
using Aspose.Email.Mapi;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/

namespace Aspose.Email.Examples.CSharp.Email
{
    class SignAMessage
    {
        public static void Run()
        {
            // ExStart:SignAMessage
            // The path to the File directory.
            string dataDir = RunExamples.GetDataDir_SMTP();
            string publicCertFile = dataDir + "MartinCertificate.cer";
            string privateCertFile = dataDir + "MartinCertificate.pfx";
            X509Certificate2 publicCert = new X509Certificate2(publicCertFile);
            X509Certificate2 privateCert = new X509Certificate2(privateCertFile, "password");
            MailMessage msg = new MailMessage("userfrom@gmail.com", "userto@gmail.com", "Signed message only", "Test Body of signed message");
            MailMessage signed = msg.AttachSignature(privateCert);
            MailMessage encrypted = signed.Encrypt(publicCert);
            MailMessage decrypted = encrypted.Decrypt(privateCert);
            MailMessage unsigned = decrypted.RemoveSignature();//The original message with proper body
            MapiMessage mapi = MapiMessage.FromMailMessage(unsigned);
            // ExEnd:SignAMessage
        }
    }
}
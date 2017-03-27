using System;
using System.Security.Cryptography.X509Certificates;
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
    class EncryptAndDecryptMessage
    {
        public static void Run()
        {
            // ExStart:EncryptAndDecryptMessage
            // The path to the File directory.
            string dataDir = RunExamples.GetDataDir_Email();

            string publicCertFile = dataDir + "MartinCertificate.cer";
            string privateCertFile = dataDir + "MartinCertificate.pfx";

            X509Certificate2 publicCert = new X509Certificate2(publicCertFile);
            X509Certificate2 privateCert = new X509Certificate2(privateCertFile, "anothertestaccount");

            // Create a message
            MailMessage msg = new MailMessage("atneostthaecrcount@gmail.com", "atneostthaecrcount@gmail.com", "Test subject", "Test Body");

            // Encrypt the message
            MailMessage eMsg = msg.Encrypt(publicCert);
            if (eMsg.IsEncrypted == true)
                Console.WriteLine("Its encrypted");
            else
                Console.WriteLine("Its NOT encrypted");

            // Decrypt the message
            MailMessage dMsg = eMsg.Decrypt(privateCert);
            if (dMsg.IsEncrypted == true)
                Console.WriteLine("Its encrypted");
            else
                Console.WriteLine("Its NOT encrypted");
            // ExEnd:EncryptAndDecryptMessage
        }
    }
}
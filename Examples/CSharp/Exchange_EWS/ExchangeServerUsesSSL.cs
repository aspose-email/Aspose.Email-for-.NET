using System;
using System.Net.Security;
using System.Security.Cryptography.X509Certificates;
using Aspose.Email.Clients.Imap;
using Aspose.Email.Clients;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/

namespace Aspose.Email.Examples.CSharp.Email.Exchange_EWS
{
    class ExchangeServerUsesSSL
    {
        // ExStart:ExchangeServerUsesSSL
        public static void Run()
        {            
            // Connect to Exchange Server using ImapClient class
            ImapClient imapClient = new ImapClient("ex07sp1", 993, "Administrator", "Evaluation1", new RemoteCertificateValidationCallback(RemoteCertificateValidationHandler));
            imapClient.SecurityOptions = SecurityOptions.SSLExplicit;

            // Select the Inbox folder
            imapClient.SelectFolder(ImapFolderInfo.InBox);

            // Get the list of messages
            ImapMessageInfoCollection msgCollection = imapClient.ListMessages();
            foreach (ImapMessageInfo msgInfo in msgCollection)
            {
                Console.WriteLine(msgInfo.Subject);
            }
            // Disconnect from the server
            imapClient.Dispose();   
        }

        // Certificate verification handler
        private static bool RemoteCertificateValidationHandler(object sender, X509Certificate certificate, X509Chain chain, SslPolicyErrors sslPolicyErrors)
        {
            return true; // ignore the checks and go ahead
        }
        // ExEnd:ExchangeServerUsesSSL
    }
}
using System;
using System.Net.Security;
using System.Security.Cryptography.X509Certificates;
using Aspose.Email.Clients.Imap;
using Aspose.Email.Clients;

namespace Aspose.Email.Examples.EWS
{
    class ExchangeServerUsesSSL
    {
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
    }
}
using Aspose.Email.Clients;
using Aspose.Email.Clients.Base;
using Aspose.Email.Clients.Imap;
using System;

namespace CSharp.IMAP
{
    public class ImapListMessagesWithMultiConnection
    {
        public static void Run()
        {
            //ExStart: 1
            ImapClient imapClient = new ImapClient();
            imapClient.Host = "<HOST>";
            imapClient.Port = 993;
            imapClient.Username = "<USERNAME>";
            imapClient.Password = "<PASSWORD>";
            imapClient.SupportedEncryption = EncryptionProtocols.Tls;
            imapClient.SecurityOptions = SecurityOptions.SSLImplicit;

            imapClient.SelectFolder("Inbox");
            imapClient.ConnectionsQuantity = 5;
            imapClient.UseMultiConnection = MultiConnectionMode.Enable;
            DateTime multiConnectionModeStartTime = DateTime.Now;
            ImapMessageInfoCollection messageInfoCol1 = imapClient.ListMessages(true);
            TimeSpan multiConnectionModeTimeSpan = DateTime.Now - multiConnectionModeStartTime;

            imapClient.UseMultiConnection = MultiConnectionMode.Disable;
            DateTime singleConnectionModeStartTime = DateTime.Now;
            ImapMessageInfoCollection messageInfoCol2 = imapClient.ListMessages(true);
            TimeSpan singleConnectionModeTimeSpan = DateTime.Now - singleConnectionModeStartTime;
            
            double performanceRelation = singleConnectionModeTimeSpan.TotalMilliseconds / multiConnectionModeTimeSpan.TotalMilliseconds;
            Console.WriteLine("Performance Relation: " + performanceRelation);
            //ExEnd: 1

            Console.WriteLine("ImapListMessagesWithMultiConnection executed successfully.");
        }
    }
}

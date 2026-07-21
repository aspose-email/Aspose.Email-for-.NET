using System;
using System.Threading;
using Aspose.Email.Clients.Imap;
using Aspose.Email.Mime;

namespace Aspose.Email.Examples.IMAP
{
    class RetrievingMessagesAsynchronously
    {
        public static void Run()
        {
            // // Connect and log in to IMAP
            // using (ImapClient client = new ImapClient("host", "username", "password"))
            // {
            //     client.SelectFolder("Issues/SubFolder");
            //     ImapMessageInfoCollection messages = client.ListMessages();
            //     AutoResetEvent evnt = new AutoResetEvent(false);
            //     MailMessage message = null;
            //     AsyncCallback callback = delegate(IAsyncResult ar)
            //     {
            //         message = client.EndFetchMessage(ar);
            //         evnt.Set();
            //     };
            //     client.BeginFetchMessage(messages[0].SequenceNumber, callback, null);
            //     evnt.WaitOne();               
            // }
        }
    }
}
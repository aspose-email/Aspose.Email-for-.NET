using System;
using Aspose.Email.Clients.Pop3;
using System.Threading;
using Aspose.Email.Mime;
using Aspose.Email.Clients;

namespace Aspose.Email.Examples.POP3
{
    class RetrieveMessagesAsynchronously
    {
        public static void Run()
        {
            // Pop3Client client = new Pop3Client();
            // client.Host = "pop.gmail.com";
            // client.Port = 995;
            // client.SecurityOptions = SecurityOptions.SSLImplicit;
            // client.Username = "username";
            // client.Password = "password";
            //
            // try
            // {
            //     Pop3MessageInfoCollection messages = client.ListMessages();
            //     Console.WriteLine("Total Number of Messages in inbox:" + messages.Count);
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
            // catch (Exception ex)
            // {
            //     Console.WriteLine(ex.Message);
            // }
        }
    }
}
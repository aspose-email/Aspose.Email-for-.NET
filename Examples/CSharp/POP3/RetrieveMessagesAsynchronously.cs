using System;
using Aspose.Email.Clients.Pop3;
using System.Threading;
using Aspose.Email.Mime;
using Aspose.Email.Clients;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/

namespace Aspose.Email.Examples.CSharp.Email.POP3
{
    class RetrieveMessagesAsynchronously
    {
        public static void Run()
        {
            // ExStart:RetrieveMessagesAsynchronously
            Pop3Client client = new Pop3Client();
            client.Host = "pop.gmail.com";
            client.Port = 995;
            client.SecurityOptions = SecurityOptions.SSLImplicit;
            client.Username = "username";
            client.Password = "password";
            
            try
            {
                Pop3MessageInfoCollection messages = client.ListMessages();
                Console.WriteLine("Total Number of Messages in inbox:" + messages.Count);
                AutoResetEvent evnt = new AutoResetEvent(false);
                MailMessage message = null;
                AsyncCallback callback = delegate(IAsyncResult ar)
                {
                    message = client.EndFetchMessage(ar);
                    evnt.Set();
                };
                client.BeginFetchMessage(messages[0].SequenceNumber, callback, null);
                evnt.WaitOne();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            // ExEnd:RetrieveMessagesAsynchronously
        }
    }
}
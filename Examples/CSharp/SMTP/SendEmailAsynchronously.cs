using System;
using Aspose.Email.Mime;
using Aspose.Email.Clients.Smtp;
using Aspose.Email.Clients;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/

namespace Aspose.Email.Examples.CSharp.Email.SMTP
{
    class SendEmailAsynchronously
    {
        // ExStart:SendEmailAsynchronously
        public static void Run()
        {
            SendMail();
        }

        static SmtpClient GetSmtpClient2()
        {
            SmtpClient client = new SmtpClient();
            client.Host = "smtp.gmail.com";
            //Specify your mail Username, Password, Port # and security option
            client.Username = "user";
            client.Password = "password";
            client.Port = 587;
            client.SecurityOptions = SecurityOptions.SSLExplicit;
            return client;
        }
        static void SendMail()
        {
            try
            {

                // Declare msg as MailMessage instance
                MailMessage msg = new MailMessage("sender@gmail.com", "receiver@gmail.com", "Test subject", "Test body");
                SmtpClient client = GetSmtpClient2();
                object state = new object();
                IAsyncResult ar = client.BeginSend(msg, Callback, state);

                Console.WriteLine("Sending message... press c to cancel mail. Press any other key to exit.");
                string answer = Console.ReadLine();

                // If the user canceled the send, and mail hasn't been sent yet,
                if (answer != null && answer.StartsWith("c"))
                {
                    client.CancelAsyncOperation(ar);
                }

                msg.Dispose();
                Console.WriteLine("Goodbye.");
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
        static AsyncCallback Callback = delegate(IAsyncResult ar)
        {
            var task = ar as IAsyncResultExt;
            if (task != null && task.IsCanceled)
            {
                Console.WriteLine("Send canceled.");
            }

            if (task != null && task.ErrorInfo != null)
            {
                Console.WriteLine("{0}", task.ErrorInfo);
            }
            else
            {
                Console.WriteLine("Message Sent.");
            }
        };
        // ExEnd:SendEmailAsynchronously
    }
}

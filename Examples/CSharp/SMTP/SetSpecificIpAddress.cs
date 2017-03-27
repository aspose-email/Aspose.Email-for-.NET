using System.Net;
using Aspose.Email.Mime;
using Aspose.Email.Clients.Smtp;
using Aspose.Email.Clients;

/* This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET 
   API reference when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq 
   for more information. If you do not wish to use NuGet, you can manually download 
   Aspose.Email for .NET API from http://www.aspose.com/downloads, 
   install it and then add its reference to this project. For any issues, questions or suggestions 
   please feel free to contact us using http://www.aspose.com/community/forums/default.aspx            
*/

namespace Aspose.Email.Examples.CSharp.Email.IMAP
{
    class SetSpecificIpAddress
    {
        // ExStart:SetSpecificIpAddress
        public static void Run()
        {
            using (SmtpClient client = new SmtpClient("smtp.gmail.com", 587, "your.email@gmail.com", "your.Password", SecurityOptions.Auto))
            {
                // set callback
                client.BindIPEndPoint += BindIPEndPointCallback;
                client.Noop();
            }
        }

        // get local endpoint callbak
        private static IPEndPoint BindIPEndPointCallback(IPEndPoint remoteEndPoint)
        {
            return new IPEndPoint(IPAddress.Any, 0);
        }
        // ExEnd:SetSpecificIpAddress
    }
}

/* This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Slides for .NET 
    * API reference when the project is build. Please check https://docs.nuget.org/consume/nuget-faq 
    * for more information. If you do not wish to use NuGet, you can manually download 
    * Aspose.Slides for .NET API from http://www.aspose.com/downloads, 
    * install it and then add its reference to this project. For any issues, questions or suggestions 
    * please feel free to contact us using http://www.aspose.com/community/forums/default.aspx            
    */

using System.Net;
using Aspose.Email;
using Aspose.Email.Mail;

namespace CSharp.SMTP
{
    class SetSpecificIpAddress
    {
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
    }
}

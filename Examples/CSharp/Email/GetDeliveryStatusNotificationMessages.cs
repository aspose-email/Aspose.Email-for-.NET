using System;
using Aspose.Email.Mime;
using Aspose.Email.Bounce;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/

namespace Aspose.Email.Examples.CSharp.Email
{
    class GetDeliveryStatusNotificationMessages
    {
        public static void Run()
        {
            // ExStart:GetDeliveryStatusNotificationMessages
            string fileName = RunExamples.GetDataDir_Email() + "failed1.msg";
            MailMessage mail = MailMessage.Load(fileName);
            BounceResult result = mail.CheckBounced();
            Console.WriteLine(fileName);
            Console.WriteLine("IsBounced : " + result.IsBounced);
            Console.WriteLine("Action : " + result.Action);
            Console.WriteLine("Recipient : " + result.Recipient);
            Console.WriteLine();
            Console.WriteLine("Reason : " + result.Reason);
            Console.WriteLine("Status : " + result.Status);
            Console.WriteLine("OriginalMessage ToAddress 1: " + result.OriginalMessage.To[0].Address);
            Console.WriteLine();
            // ExEnd:GetDeliveryStatusNotificationMessages
        }
    }
}

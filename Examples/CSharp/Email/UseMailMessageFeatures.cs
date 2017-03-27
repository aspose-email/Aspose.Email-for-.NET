using Aspose.Email.Mime;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/

namespace Aspose.Email.Examples.CSharp.Email
{
    class UseMailMessageFeatures
    {
        public static void Run()
        {
            // The path to the File directory.
            string dataDir = RunExamples.GetDataDir_Email();

            // ExStart:MailMessageFeatures
            // Create a new message
            MailMessage message = new MailMessage();
            message.From = "sender@gmail.com";
            message.To = "receiver@gmail.com";
            message.Subject = "Using MailMessage Features";

            // Specify message date
            message.Date = DateTime.Now;

            // Specify message priority
            message.Priority = MailPriority.High;

            // Specify message sensitivity
            message.Sensitivity = MailSensitivity.Normal;

            // Specify options for delivery notifications
            message.DeliveryNotificationOptions = DeliveryNotificationOptions.OnSuccess;
            // ExEnd:MailMessageFeatures
            
            message.Save(dataDir + "UseMailMessageFeatures_out.eml", SaveOptions.DefaultEml);
        }
    }
}

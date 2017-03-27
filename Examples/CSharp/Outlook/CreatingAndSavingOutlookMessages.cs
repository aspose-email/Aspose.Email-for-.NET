using Aspose.Email.Mime;
using Aspose.Email.Mapi;
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

namespace Aspose.Email.Examples.CSharp.Email.Outlook
{
    class CreatingAndSavingOutlookMessages
    {
        public static void Run()
        {
            string dataDir = RunExamples.GetDataDir_Outlook();

            // ExStart:CreatingAndSavingOutlookMessages
            // Create an instance of the MailMessage class
            MailMessage mailMsg = new MailMessage();

            // Set from, to, subject and body properties
            mailMsg.From = "sender@domain.com";
            mailMsg.To = "receiver@domain.com";
            mailMsg.Subject = "This is test message";
            mailMsg.Body = "This is test body";

            // Create an instance of the MapiMessage class and pass MailMessage as argument
            MapiMessage outlookMsg = MapiMessage.FromMailMessage(mailMsg);
            
            // Save the message (MSG) file
            string strMsgFile = @"CreatingAndSavingOutlookMessages_out.msg";
            outlookMsg.Save(dataDir + strMsgFile);
            // ExEnd:CreatingAndSavingOutlookMessages
        }
    }
}

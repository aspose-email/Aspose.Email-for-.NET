using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Aspose.Email;
using Aspose.Email.Examples.CSharp.Email;
using Aspose.Email.Mime;
using Aspose.Email.Mapi;
using Aspose.Email.Storage.Pst;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/

namespace Aspose.Email.Examples.CSharp.Email.Outlook
{
    class CreatEMLFileAndConvertToMSG
    {
        public static void Run()
        {
            // ExStart:CreatEMLFileAndConvertToMSG
            // The path to the File directory.
            string dataDir = RunExamples.GetDataDir_Outlook();

            // Create new message and save as EML
            MailMessage message = new MailMessage("from@domain.com", "to@domain.com");
            message.Subject = "subject of email";
            message.HtmlBody = "<b>Eml to msg conversion using Aspose.Email</b>" +
                "<br><hr><br><font color=blue>This is a test eml file which will be converted to msg format.</font>";
            // Add attachments
            message.Attachments.Add(new Attachment(dataDir + "attachment_1.doc"));
            message.Attachments.Add(new Attachment(dataDir + "download.png"));

            // Save it EML
            message.Save(dataDir + "CreatEMLFileAndConvertToMSG_out.eml", SaveOptions.DefaultEml);
            // Save it to MSG
            message.Save(dataDir + "CreatEMLFileAndConvertToMSG_out.msg", SaveOptions.DefaultMsgUnicode);
            // ExEnd:CreatEMLFileAndConvertToMSG
        }
    }
}
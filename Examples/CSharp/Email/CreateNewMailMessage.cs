using Aspose.Email.Mime;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/

namespace Aspose.Email.Examples.CSharp.Email
{
    class CreateNewMailMessage
    {
        public static void Run()
        {
            // ExStart:CreateNewMailMessage
            // The path to the File directory.
            string dataDir = RunExamples.GetDataDir_Email();

            // Create a new instance of MailMessage class
            MailMessage message = new MailMessage();

            // Set subject of the message, Html body and sender information
            message.Subject = "New message created by Aspose.Email for .NET";
            message.HtmlBody = "<b>This line is in bold.</b> <br/> <br/>" + "<font color=blue>This line is in blue color</font>";
            message.From = new MailAddress("from@domain.com", "Sender Name", false);

            // Add TO recipients and Add CC recipients
            message.To.Add(new MailAddress("to1@domain.com", "Recipient 1", false));
            message.To.Add(new MailAddress("to2@domain.com", "Recipient 2", false));
            message.CC.Add(new MailAddress("cc1@domain.com", "Recipient 3", false));
            message.CC.Add(new MailAddress("cc2@domain.com", "Recipient 4", false));

            // Save message in EML, EMLX, MSG and MHTML formats
            message.Save(dataDir + "CreateNewMailMessage_out.eml", SaveOptions.DefaultEml);
            message.Save(dataDir + "CreateNewMailMessage_out.emlx", SaveOptions.CreateSaveOptions(MailMessageSaveType.EmlxFormat));
            message.Save(dataDir + "CreateNewMailMessage_out.msg", SaveOptions.DefaultMsgUnicode);
            message.Save(dataDir + "CreateNewMailMessage_out.mhtml", SaveOptions.DefaultMhtml);
            // ExEnd:CreateNewMailMessage
        }
    }
}
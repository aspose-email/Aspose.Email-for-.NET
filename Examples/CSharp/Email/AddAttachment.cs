using Aspose.Email.Examples.CSharp.Email;
using Aspose.Email.Mail;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/
namespace CSharp.Email
{
    class AddAttachment
    {
        public static void Run()
        {
            // ExStart:AddAttachment
            // The path to the File directory.
            string dataDir = RunExamples.GetDataDir_Email();

            // Create an instance of MailMessage class
            MailMessage message = new MailMessage { From = "sender@sender.com" };
            message.To.Add("receiver@gmail.com");

            // Load an attachment
            Attachment attachment = new Attachment(dataDir + "1.txt");

            // Add Multiple Attachment in instance of MailMessage class and Save message to disk
            message.Attachments.Add(attachment);
            message.AddAttachment(new Attachment(dataDir + "1.jpg"));
            message.AddAttachment(new Attachment(dataDir + "1.doc"));
            message.AddAttachment(new Attachment(dataDir + "1.rar"));
            message.AddAttachment(new Attachment(dataDir + "1.pdf"));
            message.Save(dataDir + "outputAttachments_out.msg", SaveOptions.DefaultMsgUnicode);
            // ExEnd:AddAttachment
        }
    }
}

using Aspose.Email.Mime;
using Aspose.Email.Mapi;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/

namespace Aspose.Email.Examples.CSharp.Email.Outlook
{
    class ConvertMSGToMIMEMessage
    {
        public static void Run()
        {
            //ExStart:ConvertMSGToMIMEMessage
            MapiMessage msg = new MapiMessage("sender@test.com","recipient1@test.com; recipient2@test.com","Test Subject","This is a body of message.");
            var options = new MailConversionOptions();
            options.ConvertAsTnef = true;
            MailMessage mail = msg.ToMailMessage(options);
            //ExEnd:ConvertMSGToMIMEMessage
        }
    }
}
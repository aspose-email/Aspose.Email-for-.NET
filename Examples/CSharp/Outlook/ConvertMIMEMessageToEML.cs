using Aspose.Email.Mime;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/

namespace Aspose.Email.Examples.CSharp.Email.Outlook
{
    class ConvertMIMEMessageToEML
    {
        public static void Run()
        {
            // ExStart:ConvertMIMEMessageToEML
            // The path to the File directory.
            string dataDir = RunExamples.GetDataDir_Outlook();
            MailMessage msg = MailMessage.Load(dataDir + "Message2.msg");
            // Save File to disk
            msg.Save(dataDir + "ConvertMIMEMessageToEML_out.eml");
            // ExEnd:ConvertMIMEMessageToEML
        }
    }
}
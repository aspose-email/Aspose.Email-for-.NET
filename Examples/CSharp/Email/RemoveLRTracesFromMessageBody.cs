/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/
namespace Aspose.Email.Examples.CSharp.Email
{
    class RemoveLRTracesFromMessageBody
    {
        public static void Run()
        {
            // ExStart:RemoveLRTracesFromMessageBody
            string dataDir = RunExamples.GetDataDir_Email();

            //sample input file
            string fileName = "EmlWithLinkedResources.eml";

            //Load the test message with Linked Resources
            MailMessage msg = MailMessage.Load(dataDir + fileName);

            //Remove a LinkedResource
            msg.LinkedResources.RemoveAt(0, true);

            //Now clear the Alternate View for linked Resources
            msg.AlternateViews[0].LinkedResources.Clear(true);
            // ExEnd:RemoveLRTracesFromMessageBody
        }
    }
}

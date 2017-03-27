using System;
using Aspose.Email.Mapi;
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
    class GetTheTextAndRTFBodies
    {
        public static void Run()
        {
            // ExStart:GetTheTextAndRTFBodies
            string dataDir = RunExamples.GetDataDir_Outlook();

            // Load mail message
            MapiMessage msg = MapiMessage.FromMailMessage(dataDir + "Message.eml");

            MapiMessageItemBase itemBase = new MapiMessage();
            // Text body
            if (itemBase.Body != null)
                Console.WriteLine(msg.Body);
            else
                Console.WriteLine("There's no text body.");

            // RTF body
            if (itemBase.BodyRtf != null)
                Console.WriteLine(msg.BodyRtf);
            else
                Console.WriteLine("There's no RTF body.");
            // ExEnd:GetTheTextAndRTFBodies
        }

    }
}
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
    class RemovePropertiesFromMSGAndAttachments
    {
        public static void Run()
        {
            // ExStart:RemovePropertiesFromMSGAndAttachments
            // The path to the File directory.
            string dataDir = RunExamples.GetDataDir_Outlook();

            MapiMessage mapi = new MapiMessage("from@doamin.com", "to@domain.com", "subject", "body");
            mapi.SetBodyContent("<html><body><h1>This is the body content</h1></body></html>", BodyContentType.Html);
            MapiMessage attachment = MapiMessage.FromFile(dataDir + @"message.msg");
            mapi.Attachments.Add("Outlook2 Test subject.msg", attachment);
            Console.WriteLine("Before removal = " + mapi.Attachments[mapi.Attachments.Count - 1].Properties.Count);

            mapi.Attachments[mapi.Attachments.Count - 1].RemoveProperty(923467779);// Delete anyone property
            Console.WriteLine("After removal = " + mapi.Attachments[mapi.Attachments.Count - 1].Properties.Count);
            mapi.Save(@"EMAIL_589265.msg");

            MapiMessage mapi2 = MapiMessage.FromFile(@"EMAIL_589265.msg");
            Console.WriteLine("Reloaded = " + mapi2.Attachments[mapi2.Attachments.Count - 1].Properties.Count);
            // ExEnd:RemovePropertiesFromMSGAndAttachments
        }
    }
}
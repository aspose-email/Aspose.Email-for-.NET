using System;
using System.IO;
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
    class ExtractEmbeddedObjects
    {
        public static void Run()
        {
            // ExStart:ExtractEmbeddedObjects
            // The path to the File directory.
            string dataDir = RunExamples.GetDataDir_Email();

            // Create an instance of MailMessage and load an email file
            MailMessage mailMsg = MailMessage.Load(dataDir + "EmailWithAttandEmbedded.eml");

            // Extract Attachments from the message
            foreach (Attachment attachment in mailMsg.Attachments)
            {
                // To display the the attachment file name
                Console.WriteLine(attachment.Name);

                // Save the attachment to disc
                attachment.Save(dataDir + attachment.Name);

                // You can also save the attachment to memory stream
                MemoryStream ms = new MemoryStream();

                attachment.Save(ms);
            }

            // Extract Inline images from the message
            foreach (LinkedResource lr in mailMsg.LinkedResources)
            {
                Console.WriteLine(lr.ContentType.Name);

                lr.Save(dataDir + lr.ContentType.Name);
            }
            // ExEnd:ExtractEmbeddedObjects
        }
    }
}

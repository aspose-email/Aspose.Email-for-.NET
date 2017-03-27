using System;
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
    class LoadMSGFiles
    {

        public static void Run()
        {
            // ExStart:LoadMSGFiles
            // The path to the File directory.
            string dataDir = RunExamples.GetDataDir_Outlook();

            // Create an instance of MapiMessage from file
            MapiMessage msg = MapiMessage.FromFile(dataDir + @"message.msg");

            // Get subject
            Console.WriteLine("Subject:" + msg.Subject);

            // Get from address
            Console.WriteLine("From:" + msg.SenderEmailAddress);

            // Get body
            Console.WriteLine("Body" + msg.Body);

            // Get recipients information
            Console.WriteLine("Recipient: " + msg.Recipients);

            // Get attachments
            foreach (MapiAttachment att in msg.Attachments)
            {
                Console.Write("Attachment Name: " + att.FileName);
                Console.Write("Attachment Display Name: " + att.DisplayName);
            }
            // ExEnd:LoadMSGFiles
        }
    }
}
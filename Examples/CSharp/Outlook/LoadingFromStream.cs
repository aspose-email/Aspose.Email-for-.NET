using System.IO;
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
    class LoadingFromStream
    {
        public static void Run()
        {
            // The path to the File directory.
            string dataDir = RunExamples.GetDataDir_Outlook();
            string dst = dataDir + "PersonalStorage.pst";

            // ExStart:LoadingFromStream
            // Create an instance of MapiMessage from file
            byte[] bytes = File.ReadAllBytes(dataDir + @"message.msg");

            using (MemoryStream stream = new MemoryStream(bytes))
            {
                stream.Seek(0, SeekOrigin.Begin);
                // Create an instance of MapiMessage from file
                MapiMessage msg = MapiMessage.FromStream(stream);

                // Get subject
                Console.WriteLine("Subject:" + msg.Subject);

                // Get from address
                Console.WriteLine("From:" + msg.SenderEmailAddress);

                // Get body
                Console.WriteLine("Body" + msg.Body);

            }
            // ExEnd:LoadingFromStream
        }
    }
}

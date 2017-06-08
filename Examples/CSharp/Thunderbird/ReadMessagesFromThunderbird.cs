using System;
using System.IO;
using Aspose.Email.Mime;
using Aspose.Email.Storage.Mbox;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/

namespace Aspose.Email.Examples.CSharp.Email
{
    class ReadMessagesFromThunderbird
    {
        public static void Run()
        {
            // ExStart:ReadMessagesFromThunderbird

            // The path to the File directory.
            string dataDir = RunExamples.GetDataDir_Thunderbird();

            // Open the storage file with FileStream
            FileStream stream = new FileStream(dataDir + "ExampleMbox.mbox", FileMode.Open, FileAccess.Read);
            // Create an instance of the MboxrdStorageReader class and pass the stream
            MboxrdStorageReader reader = new MboxrdStorageReader(stream, false);
            // Start reading messages
            MailMessage message = reader.ReadNextMessage();

            // Read all messages in a loop
            while (message != null)
            {
                // Manipulate message - show contents
                Console.WriteLine("Subject: " + message.Subject);
                // Save this message in EML or MSG format
                message.Save(message.Subject + ".eml", SaveOptions.DefaultEml);
                message.Save(message.Subject + ".msg", SaveOptions.DefaultMsgUnicode);

                // Get the next message
                message = reader.ReadNextMessage();
            }
            // Close the streams
            reader.Dispose();
            stream.Close();
            // ExEnd:ReadMessagesFromThunderbird
        }
    }
}
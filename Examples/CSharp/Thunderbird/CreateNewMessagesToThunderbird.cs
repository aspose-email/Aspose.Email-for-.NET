using System;
using System.IO;
using Aspose.Email.Formats.Mbox;
using Aspose.Email.Mail;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/

namespace Aspose.Email.Examples.CSharp.Email
{
    class CreateNewMessagesToThunderbird
    {
        public static void Run()
        {
            // ExStart:CreateNewMessagesToThunderbird
            // The path to the File directory.
            string dataDir = RunExamples.GetDataDir_SMTP();

            try
            {
                // Open the storage file with FileStream
                FileStream stream = new FileStream(dataDir + "Please add your Thunderbird file name here", FileMode.Open, FileAccess.Read);

                // Initialize MboxStorageWriter and pass the above stream to it
                MboxrdStorageWriter writer = new MboxrdStorageWriter(stream, false);
                // Prepare a new message using the MailMessage class
                MailMessage message = new MailMessage("from@domain.com", "to@domain.com", "added 1 from Aspose.Email", "added from Aspose.Email");
                // Add this message to storage
                writer.WriteMessage(message);
                // Close all related streams
                writer.Dispose();
                stream.Close();
                // ExEnd:CreateNewMessagesToThunderbird
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message + "\nPlease add Thunderbird file name to the FileStream");
            }
        }
    }
}
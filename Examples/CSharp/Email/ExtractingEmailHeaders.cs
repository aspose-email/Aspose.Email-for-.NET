using System.IO;
using System;
using Aspose.Email.Mime;
using Aspose.Email.Mapi;

namespace Aspose.Email.Examples.CSharp.Email
{
    class ExtractingEmailHeaders
    {
        public static void Run()
        {
            // The path to the File directory.
            string dataDir = RunExamples.GetDataDir_Email();

            // ExStart:ExtractingEmailHeaders
            MailMessage message;

            // Create MailMessage instance by loading an EML file
            message = MailMessage.Load(dataDir + "email-headers.eml", new EmlLoadOptions());
            Console.WriteLine("\n\nheaders:\n\n");

            // Print out all the headers
            int index = 0;
            foreach (String header in message.Headers)
            {
                Console.Write(header + " - ");
                Console.WriteLine(message.Headers.Get(index++));
            }
            // ExEnd:ExtractingEmailHeaders
        }
    }
}

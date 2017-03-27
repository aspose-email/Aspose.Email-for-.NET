using Aspose.Email.Clients;
using Aspose.Email.Clients.Pop3;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/

namespace Aspose.Email.Examples.CSharp.Email.POP3
{
    class Pop3ClientActivityLogging
    {
        public static void Run()
        {
            string dataDir = RunExamples.GetDataDir_POP3();

            // ExStart:Pop3ClientActivityLogging
            Pop3Client client = new Pop3Client("pop.gmail.com", 995, "user@gmail.com", "password");

            // Set security mode
            client.SecurityOptions = SecurityOptions.Auto;

            try
            {
                // Get the message info collection
                Pop3MessageInfoCollection list = client.ListMessages();

                // Download each message
                for (int i = 0; i < list.Count; i++)
                {
                    // Save the EML file locally
                    client.SaveMessage(list[i].UniqueId, dataDir + list[i].UniqueId + ".eml");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            // ExEnd:Pop3ClientActivityLogging
        }
    }
}

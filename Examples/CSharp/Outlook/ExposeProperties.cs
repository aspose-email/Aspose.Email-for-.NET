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
    class ExposeProperties
    {
        public static void Run()
        {
             string dataDir = RunExamples.GetDataDir_Outlook();

            // Load mail message
            MailMessage msg = MailMessage.Load(dataDir + "Message.eml", new EmlLoadOptions());
          
            // Subject
            if (msg.Subject != null)
                Console.WriteLine(msg.Subject);
            else
                Console.WriteLine("no subject");

               // From
            if (msg.From != null)
                Console.WriteLine(msg.From);
            else
                Console.WriteLine("No sender");

            // To
            if (msg.To != null)
                Console.WriteLine(msg.To);
            else
                Console.WriteLine("No one in To");

            // Cc
            if (msg.CC != null)
                Console.WriteLine(msg.CC);
            else
                Console.WriteLine("No one in cc");
        }
    }
}
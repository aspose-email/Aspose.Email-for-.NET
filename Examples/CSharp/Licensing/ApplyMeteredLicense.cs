using Aspose.Email;
using Aspose.Email.Examples.CSharp.Email;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace CSharp.Licensing
{
    class ApplyMeteredLicense
    {
        public static void Run()
        {
            //ExStart: ApplyMeteredLicense
            Aspose.Email.Metered metered = new Aspose.Email.Metered();
            // Access the SetMeteredKey property and pass public and private keys as parameters
            metered.SetMeteredKey("*****", "*****");

            // The path to the documents directory. 
            string dataDir = RunExamples.GetDataDir_Email();

            // Load the document from disk.
            MailMessage eml = MailMessage.Load(dataDir + "Message.eml");
            //Get the page count of document
            Console.WriteLine(eml.Subject);
            //ExEnd: ApplyMeteredLicense
        }
    }
}

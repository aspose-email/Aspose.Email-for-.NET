using System;
using System.Text;
using Aspose.Email.Mime;
using Aspose.Email.Clients.Exchange.WebService;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/


namespace Aspose.Email.Examples.CSharp.Email.Exchange
{
    class SendExchangeTask
    {
        public static void Run()
        {

            try
            {

                // The path to the File directory.
                string dataDir = RunExamples.GetDataDir_Exchange();
                string dstEmail = dataDir + "Message.eml";

                // Create instance of ExchangeClient class by giving credentials
                IEWSClient client = EWSClient.GetEWSClient("https://outlook.office365.com/ews/exchange.asmx", "testUser", "pwd", "domain");

                // load task from .eml file
                EmlLoadOptions loadOptions = new EmlLoadOptions();

                loadOptions.PrefferedTextEncoding = Encoding.UTF8;
                loadOptions.PreserveTnefAttachments = true;

                // load task from .msg file
                MailMessage eml = MailMessage.Load(dstEmail, loadOptions);
                eml.From = "firstname.lastname@domain.com";
                eml.To.Clear();
                eml.To.Add(new MailAddress("firstname.lastname@domain.com"));
                client.Send(eml);
                Console.WriteLine(Environment.NewLine + "Task sent on Exchange Server successfully.");
            }
            catch (Exception ex)
            {

                Console.Write(ex.Message);
            }
        }
    }
}

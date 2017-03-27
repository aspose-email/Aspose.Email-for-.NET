using System;
using System.Net;
using Aspose.Email.Mime;
using Aspose.Email.Clients.Exchange.WebService;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/

namespace Aspose.Email.Examples.CSharp.Email.Exchange_EWS
{
    class SendTaskRequestUsingIEWSClient
    {
        public static void Run()
        {
            try
            {
                // ExStart:SendTaskRequestUsingIEWSClient
                string dataDir = RunExamples.GetDataDir_Exchange();
                // Create instance of ExchangeClient class by giving credentials
                IEWSClient client = EWSClient.GetEWSClient("https://outlook.office365.com/ews/exchange.asmx", "testUser", "pwd", "domain");

                MsgLoadOptions options = new MsgLoadOptions();
                options.PreserveTnefAttachments = true;

                // load task from .msg file
                MailMessage eml = MailMessage.Load(dataDir + "task.msg", options);
                eml.From = "firstname.lastname@domain.com";
                eml.To.Clear();
                eml.To.Add(new MailAddress("firstname.lastname@domain.com"));
                client.Send(eml);
                // ExEnd:SendTaskRequestUsingIEWSClient
            }
            catch (Exception ex)
            {
                Console.Write(ex.Message);
            }
        }
    }
}
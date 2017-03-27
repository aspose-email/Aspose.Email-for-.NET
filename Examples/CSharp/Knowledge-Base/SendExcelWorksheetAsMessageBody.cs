using Aspose.Cells;
using Aspose.Email.Clients;
using Aspose.Email.Clients.Smtp;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/

namespace Aspose.Email.Examples.CSharp.Email.Knowledge.Base
{
    class SendExcelWorksheetAsMessageBody
    {
        public static void Run()
        {
            try
            {
                // ExStart:SendExcelWorksheetAsMessageBody
                // The path to the File directory
                string dataDir = RunExamples.GetDataDir_KnowledgeBase();

                // Load the desired workbook from disk
                Workbook workbook = new Workbook(dataDir + "Data.xlsx");

                // Save the workbook to Memory Stream in HTML format
                MemoryStream ms = new MemoryStream();
                workbook.Save(ms, SaveFormat.Html);
                ms.Position = 0;

                // Define a StreamReader object with the above MemoryStream
                StreamReader sr = new StreamReader(ms);

                // Load the saved HTML from StreamReader now into a string variable
                string strHtmlBody = sr.ReadToEnd();

                // Define a new Message object and set its HtmlBody
                MailMessage message = new MailMessage();
                message.HtmlBody = strHtmlBody;
                message.Subject = "Inline Excel Message";
                message.From = "sender@abc.com";
                message.To = "receiver@xyz.com";
                message.IsBodyHtml = true;
                SmtpClient client = new SmtpClient();
                client.Host = "smtp.gmail.com";
                client.Username = "Username";
                client.Password = "Password";
                client.Port = 587;
                client.SecurityOptions = SecurityOptions.Auto;
                client.Send(message);
                // ExEnd:SendExcelWorksheetAsMessageBody
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
    }
}

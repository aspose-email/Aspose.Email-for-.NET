using Aspose.Email.Clients;
using Aspose.Email.Clients.Smtp;
using Aspose.Email.Mime;
using Aspose.Email.Tools.Merging;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.Text;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/

namespace Aspose.Email.Examples.CSharp.Email.SMTP
{
    class RowWiseMailMerge
    {        
        public static void Run()
        {
            // The path to the File directory.
            string dataDir = RunExamples.GetDataDir_SMTP();
            string dstEmail = dataDir + "EmbeddedImage.msg";

            // Create a new MailMessage instance
            MailMessage msg = new MailMessage();

            // Add subject and from address
            msg.Subject = "Hello, #FirstName#";
            msg.From = "sender@sender.com";

            // Add email address to send email also Add mesage field to HTML body
            msg.To.Add("your.email@gmail.com");
            msg.HtmlBody = "Your message here";
            msg.HtmlBody += "Thank you for your interest in <STRONG>Aspose.Email</STRONG>.";

            // Use GetSignment as the template routine, which will provide the same signature
            msg.HtmlBody += "<br><br>Have fun with it.<br><br>#GetSignature()#";

            // Create a new TemplateEngine with the MSG message,  Register GetSignature routine. It will be used in MSG.
            TemplateEngine engine = new TemplateEngine(msg);
            engine.RegisterRoutine("GetSignature", GetSignature);

            // Create an instance of DataTable and Fill a DataTable as data source
            DataTable dt = new DataTable();
            dt.Columns.Add("Receipt", typeof(string));
            dt.Columns.Add("FirstName", typeof(string));
            dt.Columns.Add("LastName", typeof(string));

            DataRow dr = dt.NewRow();
            dr["Receipt"] = "abc<asposetest123@gmail.com>";
            dr["FirstName"] = "a";
            dr["LastName"] = "bc";
            dt.Rows.Add(dr);
            dr = dt.NewRow();
            dr["Receipt"] = "John<email.2@gmail.com>";
            dr["FirstName"] = "John";
            dr["LastName"] = "Doe";
            dt.Rows.Add(dr);
            dr = dt.NewRow();
            dr["Receipt"] = "Third Recipient<email.3@gmail.com>";
            dr["FirstName"] = "Third";
            dr["LastName"] = "Recipient";
            dt.Rows.Add(dr);

            MailMessage message;
            try
            {
                foreach (DataRow currentRow in dt.Rows)
                {
                    // ExStart:RowWiseMailMerge
                    // Create message from the data in current row.
                    message = engine.Merge(currentRow);
                    // ExEnd:RowWiseMailMerge

                    // Create an instance of SmtpClient and specify server, port, username and password
                    SmtpClient client = new SmtpClient("smtp.gmail.com", 587, "your.email@gmail.com", "your.password");
                    client.SecurityOptions = SecurityOptions.Auto;

                    // Send messages in bulk
                    client.Send(message);
                }            
            }
            catch (MailException ex)
            {
                Debug.WriteLine(ex.ToString());
            }

            catch (SmtpException ex)
            {
                Debug.WriteLine(ex.ToString());
            }

            Console.WriteLine(Environment.NewLine + "Message sent after performing mail merge.");
        }

        // Template routine to provide signature
        static object GetSignature(object[] args)
        {
            return "Aspose.Email Team<br>Aspose Ltd.<br>" + DateTime.Now.ToShortDateString();
        }
    }
}

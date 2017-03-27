using System;
using System.Data;
using System.Diagnostics;
using Aspose.Email.Mime;
using Aspose.Email.Tools.Merging;
using Aspose.Email.Clients.Smtp;
using Aspose.Email.Clients;

namespace Aspose.Email.Examples.CSharp.Email.IMAP
{
    class MailMerge
    {
        // ExStart:MailMerge
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

            MailMessageCollection messages;
            try
            {
                // Create messages from the message and datasource.
                messages = engine.Instantiate(dt);

                // Create an instance of SmtpClient and specify server, port, username and password
                SmtpClient client = new SmtpClient("smtp.gmail.com", 587, "your.email@gmail.com", "your.password");
                client.SecurityOptions = SecurityOptions.Auto;

                // Send messages in bulk
                client.Send(messages);
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
        // ExEnd:MailMerge
    }
}

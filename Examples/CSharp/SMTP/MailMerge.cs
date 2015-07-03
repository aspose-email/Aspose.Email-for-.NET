//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2015 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Email. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System.IO;
using System;
using Aspose.Email.Mail;
using Aspose.Email.Outlook;
using Aspose.Email.Pop3;
using Aspose.Email;
using Aspose.Email.Mime;
using Aspose.Email.Imap;
using System.Configuration;
using System.Data;

namespace CSharp.SMTP
{
    class MailMerge
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_SMTP();
            string dstEmail = dataDir + "EmbeddedImage.msg";

            //Create a new MailMessage instance
            MailMessage msg = new MailMessage();

            //Add subject and from address
            msg.Subject = "Hello, #FirstName#";
            msg.From = "sender@sender.com";

            //Add email address to send email
            msg.To.Add("your.email@gmail.com");

            //Add mesage field to HTML body
            msg.HtmlBody = "Your message here";
            msg.HtmlBody += "Thank you for your interest in <STRONG>Aspose.Email</STRONG>.";

            //Use GetSignment as the template routine, which will provide the same signature
            msg.HtmlBody += "<br><br>Have fun with it.<br><br>#GetSignature()#";

            //Create a new TemplateEngine with the MSG message.
            TemplateEngine engine = new TemplateEngine(msg);

            // Register GetSignature routine. It will be used in MSG.
            engine.RegisterRoutine("GetSignature", new TemplateRoutine(GetSignature));

            //Create an instance of DataTable
            //Fill a DataTable as data source
            DataTable dt = new DataTable();
            dt.Columns.Add("Receipt", typeof(string));
            dt.Columns.Add("FirstName", typeof(string));
            dt.Columns.Add("LastName", typeof(string));

            //Create an instance of DataRow
            DataRow dr;

            dr = dt.NewRow();
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
                //Create messages from the message and datasource.
                messages = engine.Instantiate(dt);

                //Create an instance of SmtpClient and specify server, port, username and password
                SmtpClient client = GetSmtpClient();

                //Send messages in bulk
                client.BulkSend(messages);
            }
            catch (MailException ex)
            {
                System.Diagnostics.Debug.WriteLine(ex.ToString());
            }

            catch (SmtpException ex)
            {
                System.Diagnostics.Debug.WriteLine(ex.ToString());
            }

            Console.WriteLine(Environment.NewLine + "Message sent after performing mail merge.");
        }

        //Template routine to provide signature
        static object GetSignature(object[] args)
        {
            return "Aspose.Email Team<br>Aspose Ltd.<br>" + DateTime.Now.ToShortDateString();
        }

        private static SmtpClient GetSmtpClient()
        {
            SmtpClient client = new SmtpClient("smtp.gmail.com", 587, "your.email@gmail.com", "your.password");
            client.SecurityOptions = SecurityOptions.Auto;

            return client;
        }
    }
}

//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2013 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Email. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
using System.IO;

using Aspose.Email;
using Aspose.Email.Mail;
using System;
using Aspose.Email.Pop3;

namespace RetrievingEmailMessages
{
    public class Program
    {
        public static void Main(string[] args)
        {
            // The path to the documents directory.
            string dataDir = Path.GetFullPath("../../../Data/");

            //Create an instance of the Pop3Client class
            Pop3Client client = new Pop3Client();

            //Specify host, username and password for your client
            client.Host = "pop.gmail.com";

            // Set username
            client.Username = "asposetest123@gmail.com";

            // Set password
            client.Password = "F123456f";

            // Set the port to 995. This is the SSL port of POP3 server
            client.Port = 995;

            // Enable SSL
            client.SecurityOptions = SecurityOptions.Auto;

            try
            {
                int messageCount = client.GetMessageCount();
                // Create an instance of the MailMessage class
                MailMessage msg;

                for (int i = 1; i <= messageCount; i++)
                {
                    //Retrieve the message in a MailMessage class
                    msg = client.FetchMessage(i);

                    Console.WriteLine("From:" + msg.From.ToString());
                    Console.WriteLine("Subject:" + msg.Subject);
                    Console.WriteLine(msg.HtmlBody);
                }

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            finally
            {
                client.Disconnect();
            } 
        }
    }
}
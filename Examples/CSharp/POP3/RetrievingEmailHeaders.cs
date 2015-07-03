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

namespace CSharp.POP3
{
    class RetrievingEmailHeaders
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_POP3();
            string dstEmail = dataDir + "message.msg";

            //Create an instance of the Pop3Client class
            Pop3Client client = new Pop3Client();

            //Specify host, username and password for your client
            client.Host = "pop.gmail.com";

            // Set username
            client.Username = "your.username@gmail.com";

            // Set password
            client.Password = "your.password";

            // Set the port to 995. This is the SSL port of POP3 server
            client.Port = 995;

            // Enable SSL
            client.SecurityOptions = SecurityOptions.Auto;

            try
            {
                HeaderCollection headers = client.GetMessageHeaders(1);

                for (int i = 0; i < headers.Count; i++)
                {
                    // Display key and value in the header collection
                    Console.Write(headers.Keys[i]);
                    Console.Write(" : ");
                    Console.WriteLine(headers.Get(i));
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

            Console.WriteLine(Environment.NewLine + "Displayed header information from emails using POP3 ");
        }
    }
}

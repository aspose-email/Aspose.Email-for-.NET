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
    class SSLEnabledPOP3Server
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_POP3();
            string dstEmail = dataDir + "1234.eml";

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

            //Connect and login to a POP3 server
            try
            {
            }
            catch (Pop3Exception ex)
            {
                Console.Write(ex.ToString());
            }

            Console.WriteLine(Environment.NewLine + "Connecting to POP3 server using SSL.");
        }
    }
}

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

namespace ExtractingEmailHeaders
{
    public class Program
    {
        public static void Main(string[] args)
        {
            // The path to the documents directory.
            string dataDir = Path.GetFullPath("../../../Data/");

            MailMessage message;

            //Create MailMessage instance by loading an EML file
            message = MailMessage.Load(dataDir + "test.eml", MailMessageLoadOptions.DefaultEml);

            Console.WriteLine("\n\nheaders:\n\n");

            //print out all the headers
            int index = 0;
            foreach (String header in message.Headers)
            {
                Console.Write(header + " - ");
                Console.WriteLine(message.Headers.Get(index++));//.GetValues(header).Length.ToString());
            }

            System.Console.Read();
        }
    }
}
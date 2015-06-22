//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2015 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Email. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
using Aspose.Email.Mail;
using Aspose.Email.Mail.Bounce;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

// This feature is available from Aspose.Email for .NET 5.4.0 onwards
namespace ProcessBouncedMsgs
{
    class Program
    {
        static void Main(string[] args)
        {
            //Initialize license

            string dataDir = Path.GetFullPath("../../../Data/");

            string fileName = dataDir + "failed1.eml";
            MailMessage mail = MailMessage.Load(fileName);
            BounceResult result = mail.CheckBounced();
            Console.WriteLine(fileName);
            Console.WriteLine("IsBounced : " + result.IsBounced);
            Console.WriteLine("Action : " + result.Action);
            Console.WriteLine("Recipient : " + result.Recipient);
            Console.WriteLine();

            fileName = dataDir + "failedReport2.eml";
            mail = MailMessage.Load(fileName);
            result = mail.CheckBounced();
            Console.WriteLine(fileName);
            Console.WriteLine("IsBounced : " + result.IsBounced);
            Console.WriteLine("Action : " + result.Action);
            Console.WriteLine("Recipient : " + result.Recipient);
            Console.WriteLine("Reason : " + result.Reason);
            Console.WriteLine("Status : " + result.Status);
            Console.WriteLine("OriginalMessage ToAddress 1: " + result.OriginalMessage.To[0].Address);
            Console.WriteLine("OriginalMessage ToAddress 2: " + result.OriginalMessage.To[1].Address);
            Console.WriteLine();
            
            fileName = dataDir + "delayed.eml";
            mail = MailMessage.Load(fileName);
            result = mail.CheckBounced();
            Console.WriteLine(fileName);
            Console.WriteLine("IsBounced : " + result.IsBounced);
            Console.WriteLine("Action : " + result.Action);
            Console.WriteLine("Recipient : " + result.Recipient);

            Console.WriteLine("\nExecution completed.");
            Console.ReadKey();
        }
    }
}

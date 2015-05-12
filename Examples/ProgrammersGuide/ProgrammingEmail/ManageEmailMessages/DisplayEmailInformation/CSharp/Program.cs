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

namespace DisplayEmailInformation
{
    public class Program
    {
        public static void Main(string[] args)
        {
            // The path to the documents directory.
            string dataDir = Path.GetFullPath("../../../Data/");

            MailMessage message;

            //Create MailMessage instance by loading an Eml file
            message = MailMessage.Load(dataDir + "test.eml", MessageFormat.Eml);

            System.Console.Write("From:");

            //Gets the sender info
            System.Console.WriteLine(message.From);
            System.Console.Write("To:");

            //Gets the recipient info
            System.Console.WriteLine(message.To);
            System.Console.Write("Subject:");

            //Gets the subject
            System.Console.WriteLine(message.Subject);
            System.Console.WriteLine("HtmlBody:");

            //Gets the htmlbody 
            System.Console.WriteLine(message.HtmlBody);
            System.Console.WriteLine("TextBody");

            //Gets the textbody
            System.Console.WriteLine(message.TextBody);
        }
    }
}
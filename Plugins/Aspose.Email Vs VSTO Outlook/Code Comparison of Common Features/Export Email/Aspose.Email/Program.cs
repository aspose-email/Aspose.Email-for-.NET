using Aspose.Email.Mail;
using System;
using System.Collections.Generic;
using System.Text;

namespace Aspose.Email
{
    class Program
    {
        static void Main(string[] args)
        {
            string FilePath = @"E:\Aspose\Aspose VS VSTO\Sample Files\ExportEmail.msg";
            
            //Create an instance of the MailMessage class
            MailMessage msg = new MailMessage();

            //Export to MHT format
            msg.Save(FilePath, SaveOptions.DefaultMsg);
        }
    }
}

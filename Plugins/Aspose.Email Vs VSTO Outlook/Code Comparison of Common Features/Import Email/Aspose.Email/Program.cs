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
            string FilePath = @"E:\Aspose\Aspose VS VSTO\Sample Files\ImportEmail.msg";
            
            // loading with default options

            // load from msg
            MailMessage eml = MailMessage.Load(FilePath, new MsgLoadOptions());
        }
    }
}

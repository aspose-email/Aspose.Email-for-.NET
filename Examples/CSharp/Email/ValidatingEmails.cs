using Aspose.Email.Tools.Verifications;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/

namespace Aspose.Email.Examples.CSharp.Email
{
    class ValidatingEmails
    {
        public static void Run()
        {
            try
            {
                // ExStart:ValidatingEmails
                EmailValidator ev = new EmailValidator();
                ValidationResult result;
                ev.Validate("user@domain.com", out result);
                if (result.ReturnCode == ValidationResponseCode.ValidationSuccess)
                {
                    Console.WriteLine("the email address is valid.");
                }
                else
                {
                    Console.WriteLine("the mail address is invalid,for the {0}", result.Message);
                }
                // ExEnd:ValidatingEmails
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
    }
}

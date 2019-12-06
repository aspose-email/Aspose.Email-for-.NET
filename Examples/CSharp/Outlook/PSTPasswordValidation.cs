using System;
using Aspose.Email.Mapi;
using Aspose.Email.Storage.Pst;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from https://www.nuget.org/packages/Aspose.Email/, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using https://forum.aspose.com/c/email
*/

namespace Aspose.Email.Examples.CSharp.Email.Outlook
{
    class PSTPasswordValidation
    {
        public static void Run()
        {
            // ExStart:1
            // The path to the File directory.
            string dataDir = RunExamples.GetDataDir_Outlook();

            using (PersonalStorage pst = PersonalStorage.FromFile(dataDir + "passwordprotectedPST.pst"))
            {
                Console.WriteLine("The storage is password protected - " + pst.Store.IsPasswordProtected);
                Console.WriteLine("Password is valid - " + pst.Store.IsPasswordValid("Password1"));
            }
            // ExEnd:1
            Console.WriteLine("PSTPasswordValidation executed successfully");
        }
    }
}
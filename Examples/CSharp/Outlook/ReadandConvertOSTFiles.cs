using System.IO;
using System;
using System.Collections.Generic;
using Aspose.Email;
using Aspose.Email.Storage.Pst;

/* This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET 
   API reference when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq 
   for more information. If you do not wish to use NuGet, you can manually download 
   Aspose.Email for .NET API from http://www.aspose.com/downloads, 
   install it and then add its reference to this project. For any issues, questions or suggestions 
   please feel free to contact us using http://www.aspose.com/community/forums/default.aspx            
*/


namespace Aspose.Email.Examples.CSharp.Email.Outlook
{
    class ReadandConvertOSTFiles
    {
        public static void Run()
        {
            //ExStart:ReadandConvertOSTFiles
            string dataDir = RunExamples.GetDataDir_Outlook();
            // Load the Outlook file
            string path = dataDir + "PersonalStorage.pst";
            PersonalStorage pst = PersonalStorage.FromFile(path);
            // Get the Display Name of the file
            Console.WriteLine("Display Format: " + pst.Format);
            //ExEnd:ReadandConvertOSTFiles
        }
    }
}

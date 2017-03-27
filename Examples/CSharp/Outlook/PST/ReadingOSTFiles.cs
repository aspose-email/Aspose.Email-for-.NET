using Aspose.Email.Storage.Pst;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

/* This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET 
   API reference when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq 
   for more information. If you do not wish to use NuGet, you can manually download 
   Aspose.Email for .NET API from http://www.aspose.com/downloads, 
   install it and then add its reference to this project. For any issues, questions or suggestions 
   please feel free to contact us using http://www.aspose.com/community/forums/default.aspx            
*/

namespace Aspose.Email.Examples.CSharp.Email.Outlook.PST
{
    class ReadingOSTFiles
    {
        public static void Run()
        {
            string dataDir = RunExamples.GetDataDir_Outlook();

            // ExStart:ReadingOSTFiles
            PersonalStorage personalStorage = PersonalStorage.FromFile(dataDir + "PersonalStorageFile.ost");
            
            // Get the format of the file
            Console.WriteLine("File Format of OST: " + personalStorage.Format);
            // ExEnd:ReadingOSTFiles
        }
    }
}

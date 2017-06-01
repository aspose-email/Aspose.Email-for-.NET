using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using Aspose.Email.Storage.Mbox;
/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/
namespace Aspose.Email.Examples.CSharp.Email
{
    class GetNumberOfItemsFromMBox
    {
        public static void Run()
        {
            //ExStart: GetNumberOfItemsFromMBox
            // The path to the File directory.
            string dataDir = RunExamples.GetDataDir_Thunderbird();

            using (FileStream stream = new FileStream(dataDir + "ExampleMbox.mbox", FileMode.Open, FileAccess.Read))
            using (MboxrdStorageReader reader = new MboxrdStorageReader(stream, false))
            {
                Console.WriteLine("Total number of messages in Mbox file: " + reader.GetTotalItemsCount());
            }
            //ExEnd: GetNumberOfItemsFromMBox
        }
    }
}

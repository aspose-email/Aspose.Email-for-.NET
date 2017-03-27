using Aspose.Words;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/

namespace Aspose.Email.Examples.CSharp.Email.Knowledge.Base
{
    class SaveEmailAsPDF
    {
        public static void Run()
        {
            // ExStart:SaveEmailAsPDF
            // The path to the File directory
            string dataDir = RunExamples.GetDataDir_KnowledgeBase();
            MailMessage mailMsg = MailMessage.Load(dataDir + "message3.msg");
            MemoryStream ms = new MemoryStream();
            mailMsg.Save(ms, Aspose.Email.SaveOptions.DefaultMhtml);

            // create an instance of LoadOptions and set the LoadFormat to Mhtml
            var loadOptions = new Aspose.Words.LoadOptions();
            loadOptions.LoadFormat = LoadFormat.Mhtml;

            // create an instance of Document and load the MTHML from MemoryStream
            var document = new Aspose.Words.Document(ms, loadOptions);

            // create an instance of HtmlSaveOptions and set the SaveFormat to Html
            var saveOptions = new Aspose.Words.Saving.PdfSaveOptions();
            document.Save(dataDir + "SaveEmailAsPDF_out.pdf", saveOptions);
            // ExEnd:SaveEmailAsPDF
        }
    }
}

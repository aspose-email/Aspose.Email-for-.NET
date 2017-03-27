using Aspose.Email.Mime;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Aspose.Email.Examples.CSharp.Email
{
    class SaveMessageAsOFT
    {
        public static void Run()
        { 
            
            /// <summary>
            /// This exmpale shows how to save an email as Outlook Template (.OFT) file using the MailMesasge API
            /// MsgSaveOptions.SaveAsTemplate - Set to true, if need to be saved as Outlook File Template(OFT format).
            /// Available from Aspose.Email for .NET 6.4.0 onwards
            /// </summary>

            string dataDir = RunExamples.GetDataDir_Email();

            // ExStart: SaveMessageAsOFT
            using (MailMessage eml = new MailMessage("test@from.to", "test@to.to", "template subject", "Template body"))
            {
                string oftEmlFileName = "EmlAsOft_out.oft";
                MsgSaveOptions options = SaveOptions.DefaultMsgUnicode;
                options.SaveAsTemplate = true;
                eml.Save(dataDir + oftEmlFileName, options);
            }
            // ExEnd: SaveMessageAsOFT
        }
    }
}

using Aspose.Email.Mail;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Aspose.Email.Examples.CSharp.Email
{
    class SaveMessageAsOFT
    {
        static void Run()
        { 
            // ExStart: SaveMessageAsOFT
            /// <summary>
            /// This exmpale shows how to save an email as Outlook Template (.OFT) file using the MailMesasge API
            /// MsgSaveOptions.SaveAsTemplate - Set to true, if need to be saved as Outlook File Template(OFT format).
            /// Available from Aspose.Email for .NET 6.4.0 onwards
            /// </summary>

            using (MailMessage eml = new MailMessage("test@from.to", "test@to.to", "template subject", "Template body"))
            {
                string oftEmlFileName = "EmlAsOft.oft";
                MsgSaveOptions options = SaveOptions.DefaultMsgUnicode;
                options.SaveAsTemplate = true;
                eml.Save(oftEmlFileName, options);
            }

            // ExEnd: SaveMessageAsOFT
        }
    }
}

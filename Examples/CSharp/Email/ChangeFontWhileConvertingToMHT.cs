using Aspose.Email;
using Aspose.Email.Examples.CSharp.Email;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace CSharp.Email
{
    public class ChangeFontWhileConvertingToMHT
    {
        public static void Run()
        {
            //ExStart: ChangeFontWhileConvertingToMHT
            string dataDir = RunExamples.GetDataDir_Email();

            MailMessage eml = MailMessage.Load(dataDir + "Message.eml");

            MhtSaveOptions options = SaveOptions.DefaultMhtml;
            int cssInsertPos = options.CssStyles.LastIndexOf("</style>");
            options.CssStyles = options.CssStyles.Insert(cssInsertPos,
                ".PlainText{" +
                "    font-family: sans-serif;" +
                "}"
                );

            eml.Save("message.mhtml", options);
            //ExEnd: ChangeFontWhileConvertingToMHT
        }
    }
}

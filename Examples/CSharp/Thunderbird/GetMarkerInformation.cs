using Aspose.Email.Formats.Mbox;
using Aspose.Email.Mail;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace Aspose.Email.Examples.CSharp.Thunderbird
{
    class GetMarkerInformation
    {
        static void Run()
        { 
            // ExStart: GetMarkerInformation
            ///<summary>
            /// This exmaple shows how to get marker information while reading or writing a message to Mbox file.
            /// Introduced in Aspose.Email for .NET 6.4.0
            ///</summary>
            using (FileStream stream = new FileStream("inbox.dat", FileMode.Open, FileAccess.Read))
            using (MboxrdStorageReader reader = new MboxrdStorageReader(stream, false))
            {
                MailMessage msg;
                string fromMarker = null;
                while ((msg = reader.ReadNextMessage(out fromMarker)) != null)
                {
                    Console.WriteLine(fromMarker);

                    msg.Dispose();
                }
            }

            using (FileStream writeStream = new FileStream("inbox.dat", FileMode.Create, FileAccess.Write))
            using (MboxrdStorageWriter writer = new MboxrdStorageWriter(writeStream, false))
            {
                string fromMarker = null;
                MailMessage msg = MailMessage.Load("input.eml");
                writer.WriteMessage(msg, out fromMarker);

                Console.WriteLine(fromMarker);
            }
            // ExEnd: GetMarkerInformation
        }
    }
}

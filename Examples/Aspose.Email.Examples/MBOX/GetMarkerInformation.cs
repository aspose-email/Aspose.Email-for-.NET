using Aspose.Email.Storage.Mbox;
using System;
using System.IO;

namespace Aspose.Email.Examples.MBOX
{
    class GetMarkerInformation
    {
        public static void Run()
        {

            using (FileStream stream = new FileStream(Data.Mbox + "ExampleMbox.mbox", FileMode.Open, FileAccess.Read))
            using (MboxrdStorageReader reader = new MboxrdStorageReader(stream, new MboxLoadOptions()))
            {
                MailMessage msg;
                string fromMarker = null;
                while ((msg = reader.ReadNextMessage(out fromMarker)) != null)
                {
                    Console.WriteLine(fromMarker);

                    msg.Dispose();
                }
            }

            using (FileStream writeStream = new FileStream(Data.Out + "ExampleMbox.mbox", FileMode.Create, FileAccess.Write))
            using (MboxrdStorageWriter writer = new MboxrdStorageWriter(writeStream, new MboxSaveOptions()))
            {
                string fromMarker = null;
                MailMessage msg = MailMessage.Load(Data.Mbox + "EmailWithAttandEmbedded.eml");
                writer.WriteMessage(msg, out fromMarker);

                Console.WriteLine(fromMarker);
            }
        }
    }
}

using Aspose.Email.Storage.Mbox;
using System;
using System.IO;

namespace Aspose.Email.Examples.CSharp.Email
{
    class GetMarkerInformation
    {
        public static void Run()
        {
            // ExStart: GetMarkerInformation
            // The path to the File directory.
            string dataDir = RunExamples.GetDataDir_Thunderbird();

            using (FileStream stream = new FileStream(dataDir + "ExampleMbox.mbox", FileMode.Open, FileAccess.Read))
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

            using (FileStream writeStream = new FileStream(dataDir + "ExampleMbox.mbox", FileMode.Create, FileAccess.Write))
            using (MboxrdStorageWriter writer = new MboxrdStorageWriter(writeStream, false))
            {
                string fromMarker = null;
                MailMessage msg = MailMessage.Load(dataDir + "EmailWithAttandEmbedded.eml");
                writer.WriteMessage(msg, out fromMarker);

                Console.WriteLine(fromMarker);
            }
            // ExEnd: GetMarkerInformation
        }
    }
}

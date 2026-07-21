using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using Aspose.Email.Storage.Mbox;
namespace Aspose.Email.Examples.MBOX
{
    class GetNumberOfItemsFromMBox
    {
        public static void Run()
        {

            using (FileStream stream = new FileStream(Data.Mbox + "ExampleMbox.mbox", FileMode.Open, FileAccess.Read))
            using (MboxrdStorageReader reader = new MboxrdStorageReader(stream, new MboxLoadOptions()))
            {
                Console.WriteLine("Total number of messages in Mbox file: " + reader.GetTotalItemsCount());
            }
        }
    }
}

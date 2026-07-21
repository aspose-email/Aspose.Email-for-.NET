// Demonstrates how to load a MapiMessage from an MSG file and cast it to a MapiTask.

using System;
using Aspose.Email.Mapi;

namespace Aspose.Email.Examples.MAPI
{
    internal static class ReadingMapiTask
    {
        public static void Run()
        {
            MapiMessage msg = MapiMessage.Load(Data.Mapi/"MapiTask.msg");
            MapiTask task = (MapiTask)msg.ToMapiMessageItem();
            Console.WriteLine("Subject: " + task.Subject);
            Console.WriteLine("Status: " + task.State);
        }
    }
}

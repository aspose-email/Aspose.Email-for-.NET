// Demonstrates how to load a MapiMessage from a byte array via a MemoryStream.

using System;
using System.IO;
using Aspose.Email.Mapi;

namespace Aspose.Email.Examples.MAPI
{
    internal static class LoadingFromStream
    {
        public static void Run()
        {
            var bytes = File.ReadAllBytes(Data.Mapi/"message.msg");

            using (var stream = new MemoryStream(bytes))
            {
                var msg = MapiMessage.Load(stream);
                Console.WriteLine("Subject: " + msg.Subject);
                Console.WriteLine("From: " + msg.SenderEmailAddress);
                Console.WriteLine("Body: " + msg.Body);
            }
        }
    }
}

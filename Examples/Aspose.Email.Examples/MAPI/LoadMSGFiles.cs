// Demonstrates how to load an MSG file and display its subject, sender, body, and attachments.

using System;
using Aspose.Email.Mapi;

namespace Aspose.Email.Examples.MAPI
{
    internal static class LoadMsgFiles
    {
        public static void Run()
        {
            var msg = MapiMessage.Load(Data.Mapi/"message.msg");

            Console.WriteLine("Subject: " + msg.Subject);
            Console.WriteLine("From: " + msg.SenderEmailAddress);
            Console.WriteLine("Body: " + msg.Body);
            Console.WriteLine("Recipients: " + msg.Recipients);

            foreach (var att in msg.Attachments)
            {
                Console.WriteLine("Attachment Name: " + att.FileName);
                Console.WriteLine("Attachment Display Name: " + att.DisplayName);
            }
        }
    }
}

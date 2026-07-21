// Demonstrates how to load a MapiContact from an Outlook MSG file.

using System;
using Aspose.Email.Mapi;

namespace Aspose.Email.Examples.MAPI
{
    internal static class LoadingContactFromMsg
    {
        public static void Run()
        {
            var msg = MapiMessage.Load(Data.Mapi/"Contact.msg");
            var contact = (MapiContact)msg.ToMapiMessageItem();
            Console.WriteLine("Name: " + contact.NameInfo.DisplayName);
        }
    }
}

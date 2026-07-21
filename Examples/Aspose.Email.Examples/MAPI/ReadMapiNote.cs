// Demonstrates how to load a MapiMessage from an MSG file and cast it to a MapiNote.

using System;
using Aspose.Email.Mapi;

namespace Aspose.Email.Examples.MAPI
{
    internal static class ReadMapiNote
    {
        public static void Run()
        {
            MapiMessage note = MapiMessage.Load(Data.Mapi/"MapiNote.msg");
            MapiNote note2 = (MapiNote)note.ToMapiMessageItem();
            Console.WriteLine("Subject: " + note2.Subject);
            Console.WriteLine("Body: " + note2.Body);
        }
    }
}

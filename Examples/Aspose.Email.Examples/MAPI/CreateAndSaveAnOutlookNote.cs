// Demonstrates how to create a MapiNote and save it as MSG.

using System;
using Aspose.Email.Mapi;

namespace Aspose.Email.Examples.MAPI
{
    internal static class CreateAndSaveAnOutlookNote
    {
        public static void Run()
        {
            var note = new MapiNote
            {
                Subject = "Blue color note",
                Body = "This is a blue color note",
                Color = NoteColor.Blue,
                Height = 500,
                Width = 500
            };
            var outputPath = Data.Out/"MapiNote_out.msg";
            note.Save(outputPath, NoteSaveFormat.Msg);

            // Colour and size are what Outlook uses to render the sticky note.
            Console.WriteLine($"Note:  {note.Subject}");
            Console.WriteLine($"Style: {note.Color}, {note.Width}x{note.Height}");
            Console.WriteLine($"Saved to {outputPath}");
        }
    }
}

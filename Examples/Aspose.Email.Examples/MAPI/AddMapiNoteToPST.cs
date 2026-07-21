// Demonstrates how to create MAPI note items with different colors and add them to a PST notes folder.

using System;
using System.IO;
using Aspose.Email.Mapi;
using Aspose.Email.Storage.Pst;

namespace Aspose.Email.Examples.MAPI
{
    internal static class AddMapiNoteToPst
    {
        public static void Run()
        {
            var mess = MapiMessage.Load(Data.Mapi/"Note.msg");

            // Colour is set explicitly on every note: the source Note.msg already carries
            // one, so an unset note would silently keep the colour of the sample file.
            var note1 = (MapiNote)mess.ToMapiMessageItem();
            note1.Subject = "Yellow color note";
            note1.Body = "This is a yellow color note";
            note1.Color = NoteColor.Yellow;

            var note2 = (MapiNote)mess.ToMapiMessageItem();
            note2.Subject = "Pink color note";
            note2.Body = "This is a pink color note";
            note2.Color = NoteColor.Pink;

            var note3 = (MapiNote)mess.ToMapiMessageItem();
            note3.Subject = "Blue color note";
            note3.Body = "This is a blue color note";
            note3.Color = NoteColor.Blue;
            note3.Height = 500;
            note3.Width = 500;

            var path = Data.Out/"AddMapiNoteToPST_out.pst";

            if (File.Exists(path))
                File.Delete(path);

            using (var pst = PersonalStorage.Create(path, FileFormatVersion.Unicode))
            {
                // A predefined folder is what makes Outlook show these as sticky notes
                // rather than as ordinary messages.
                var notesFolder = pst.CreatePredefinedFolder("Notes", StandardIpmFolder.Notes);

                foreach (var note in new[] { note1, note2, note3 })
                {
                    notesFolder.AddMapiMessageItem(note);
                    Console.WriteLine($"{note.Subject,-20} colour: {note.Color}, {note.Width}x{note.Height}");
                }

                Console.WriteLine($"\nAdded {notesFolder.ContentCount} note(s) to {path}");
            }
        }
    }
}

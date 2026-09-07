// Demonstrates opening a PST from a stream with explicit load options, and reading
// back what the storage can actually do.
//
// LeaveStreamOpen matters when the caller owns the stream - without it, disposing the
// storage closes the stream too, which breaks anything still reading from it.

using System;
using System.IO;
using Aspose.Email.Storage.Pst;

namespace Aspose.Email.Examples.PST
{
    internal static class OpenPstFromStreamWithLoadOptions
    {
        public static void Run()
        {
            // Work on a copy: the storage is opened for writing.
            var pstPath = Data.Out/"OpenPstFromStreamWithLoadOptions_out.pst";
            File.Copy(Data.Mapi/"Sub.pst", pstPath, true);

            using (var stream = File.Open(pstPath, FileMode.Open, FileAccess.ReadWrite))
            {
                var loadOptions = new PersonalStorageLoadOptions
                {
                    Writable = true,
                    LeaveStreamOpen = true
                };

                using (var pst = PersonalStorage.FromStream(stream, loadOptions))
                {
                    Console.WriteLine($"Format:        {pst.Format}");
                    Console.WriteLine($"Unicode:       {pst.IsUnicode}");
                    Console.WriteLine($"Writable:      {pst.CanWrite}");
                    Console.WriteLine($"Store name:    {pst.Store.DisplayName}");
                    Console.WriteLine($"Items:         {pst.Store.GetTotalItemsCount()}");
                    Console.WriteLine($"Root folder:   {pst.RootFolder.DisplayName}");
                }

                // LeaveStreamOpen kept the stream usable after the storage was disposed.
                Console.WriteLine($"\nStream still open: {stream.CanRead} ({stream.Length:N0} bytes)");
            }

            // Opening read-only is the default when no options are passed.
            using (var pst = PersonalStorage.FromFile(pstPath, false))
                Console.WriteLine($"Reopened read-only, CanWrite: {pst.CanWrite}");
        }
    }
}

// Demonstrates how to load a PST/OST file and display its format.

using System;
using Aspose.Email.Storage.Pst;

namespace Aspose.Email.Examples.MAPI
{
    internal static class ReadAndConvertOstFiles
    {
        public static void Run()
        {
            using (PersonalStorage pst = PersonalStorage.FromFile(Data.Mapi/"PersonalStorage.pst"))
            {
                Console.WriteLine("Display Format: " + pst.Format);
            }
        }
    }
}

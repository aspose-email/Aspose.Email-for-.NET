// Demonstrates how to open a Lotus Notes NSF storage file and enumerate
// all messages, printing each message's subject.

using System;
using Aspose.Email.Storage.Nsf;

namespace Aspose.Email.Examples.Email
{
    internal static class ReadMessagesFromNsfStorage
    {
        public static void Run()
        {
            using (var nsf = new NotesStorageFacility(Data.Email/"SampleNSF.nsf"))
            {
                foreach (var eml in nsf.EnumerateMessages())
                    Console.WriteLine(eml.Subject);
            }
        }
    }
}

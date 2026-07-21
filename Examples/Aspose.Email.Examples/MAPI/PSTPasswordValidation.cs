// Demonstrates how to check if a PST file is password-protected and validate the password.

using System;
using Aspose.Email.Storage.Pst;

namespace Aspose.Email.Examples.MAPI
{
    internal static class PstPasswordValidation
    {
        public static void Run()
        {
            using (PersonalStorage pst = PersonalStorage.FromFile(Data.Mapi/"passwordprotectedPST.pst"))
            {
                Console.WriteLine("The storage is password protected - " + pst.Store.IsPasswordProtected);
                Console.WriteLine("Password is valid - " + pst.Store.IsPasswordValid("Password1"));
            }
        }
    }
}

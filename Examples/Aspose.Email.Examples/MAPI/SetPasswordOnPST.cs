// Demonstrates how to set and then remove the password of an Outlook PST file.

using System;
using Aspose.Email.Storage.Pst;

namespace Aspose.Email.Examples.MAPI
{
    internal static class SetPasswordOnPST
    {
        public static void Run()
        {
            var pstPath = Data.Out/"SetPasswordOnPST_out.pst";

            using (var pst = PersonalStorage.Create(pstPath, FileFormatVersion.Unicode))
            {
                Console.WriteLine($"Password protected on creation: {pst.Store.IsPasswordProtected}");

                pst.Store.ChangePassword("Password1");
                Console.WriteLine($"After setting a password:       {pst.Store.IsPasswordProtected}");
                Console.WriteLine($"Password is valid:              {pst.Store.IsPasswordValid("Password1")}");

                // Passing null clears the password again.
                pst.Store.ChangePassword(null);
                Console.WriteLine($"After removing the password:    {pst.Store.IsPasswordProtected}");
            }

            Console.WriteLine($"\nSaved to {pstPath}");
        }
    }
}

// Demonstrates how to validate an email address using EmailValidator
// and display the validation result.

using System;
using Aspose.Email.Tools.Verifications;

namespace Aspose.Email.Examples.Email
{
    internal static class ValidatingEmails
    {
        public static void Run()
        {
            try
            {
                var ev = new EmailValidator();
                ev.Validate("user@domain.com", out var result);

                if (result.ReturnCode == ValidationResponseCode.ValidationSuccess)
                    Console.WriteLine("The email address is valid.");
                else
                    Console.WriteLine($"The email address is invalid: {result.Message}");
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
    }
}

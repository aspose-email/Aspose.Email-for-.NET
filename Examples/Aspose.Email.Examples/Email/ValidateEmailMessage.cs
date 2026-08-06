// Demonstrates how to check that a message file is well formed before processing it,
// and how to read the reported problems when it is not.

using System;
using Aspose.Email.Tools.Verifications;

namespace Aspose.Email.Examples.Email
{
    internal static class ValidateEmailMessage
    {
        public static void Run()
        {
            foreach (var fileName in new[] { "Message.eml", "Message.msg", "Message.mhtml", "1.jpg" })
                Validate(Data.Email/fileName, fileName);
        }

        private static void Validate(string path, string label)
        {
            var result = MessageValidator.Validate(path);

            Console.WriteLine($"{label}: {(result.IsSuccess ? "valid" : "not valid")} (detected as {result.FormatType})");

            if (result.IsSuccess)
                return;

            if (!string.IsNullOrEmpty(result.ErrorMessage))
                Console.WriteLine($"  {result.ErrorMessage}");

            foreach (var error in result.Errors)
                Console.WriteLine($"  {error.ErrorType}: {error.Description}");
        }
    }
}

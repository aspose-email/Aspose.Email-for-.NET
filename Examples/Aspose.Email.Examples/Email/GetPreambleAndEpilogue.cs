// Demonstrates the preamble and epilogue of a multipart message - the text that sits
// before the first boundary and after the last one. Mail clients ignore it, but it
// survives a round trip through MailMessage.

using System;

namespace Aspose.Email.Examples.Email
{
    internal static class GetPreambleAndEpilogue
    {
        public static void Run()
        {
            var eml = MailMessage.Load(Data.Email/"Attachments.eml");

            Console.WriteLine($"Preamble read from the file: {Show(eml.Preamble)}");
            Console.WriteLine($"Epilogue read from the file: {Show(eml.Epilogue)}");

            eml.Preamble = "This is a multi-part message in MIME format.";
            eml.Epilogue = "End of the MIME message.";

            var outputPath = Data.Out/"GetPreambleAndEpilogue_out.eml";
            eml.Save(outputPath, SaveOptions.DefaultEml);

            var reloaded = MailMessage.Load(outputPath);
            Console.WriteLine($"\nPreamble after the round trip: {Show(reloaded.Preamble)}");
            Console.WriteLine($"Epilogue after the round trip: {Show(reloaded.Epilogue)}");
            Console.WriteLine($"\nSaved to {outputPath}");
        }

        private static string Show(string value) =>
            string.IsNullOrWhiteSpace(value) ? "(empty)" : $"'{value.Trim()}'";
    }
}

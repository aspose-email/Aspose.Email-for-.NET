using System;
using System.IO;

namespace Aspose.Email.Examples
{
    class ApiExamples
    {
        [STAThread]
        static void Main(string[] args)
        {
            ApplyLicense();
            Data.ClearOut();
            ExampleRunner.Run(args);
        }

        private static void ApplyLicense()
        {
            var licensePath = FindLicense();
            if (licensePath != null)
            {
                new License().SetLicense(licensePath);
                return;
            }

            Console.ForegroundColor = ConsoleColor.DarkYellow;
            Console.WriteLine("Note: running in evaluation mode - output will be watermarked and truncated.");
            Console.WriteLine("      To apply your license, either drop Aspose.Email.NET.lic into the");
            Console.WriteLine("      solution folder or set the ASPOSE_EMAIL_LICENSE environment variable.");
            Console.ResetColor();
            Console.WriteLine();
        }

        // Locates the license file, or returns null when running unlicensed.
        private static string FindLicense()
        {
            // An explicit environment variable always wins, for example:
            //   Windows:      set ASPOSE_EMAIL_LICENSE=C:\licenses\Aspose.Email.NET.lic
            //   macOS/Linux:  export ASPOSE_EMAIL_LICENSE=/licenses/Aspose.Email.NET.lic
            var fromEnvironment = Environment.GetEnvironmentVariable("ASPOSE_EMAIL_LICENSE");
            if (!string.IsNullOrEmpty(fromEnvironment) && File.Exists(fromEnvironment))
                return fromEnvironment;

            // Otherwise fall back to any .lic file sitting in the solution folder,
            // so that a freshly cloned copy runs licensed without extra setup.
            var root = Data.RootPath;
            if (root == null)
                return null;

            var preferred = Path.Combine(root, "Aspose.Email.NET.lic");
            if (File.Exists(preferred))
                return preferred;

            var anyLicense = Directory.GetFiles(root, "*.lic");
            return anyLicense.Length > 0 ? anyLicense[0] : null;
        }
    }
}

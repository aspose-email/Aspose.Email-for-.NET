// Demonstrates how to apply a metered license, which bills by API usage instead of
// by a license file.

using System;
using Aspose.Email;
namespace Aspose.Email.Examples.Licensing
{
    internal static class ApplyMeteredLicense
    {
        public static void Run()
        {
            // Metered keys are credentials, so they are read from the environment rather
            // than hard-coded here. Get yours from the Aspose dashboard and set:
            //   Windows:      set ASPOSE_EMAIL_METERED_PUBLIC_KEY=...
            //   macOS/Linux:  export ASPOSE_EMAIL_METERED_PUBLIC_KEY=...
            var publicKey = Environment.GetEnvironmentVariable("ASPOSE_EMAIL_METERED_PUBLIC_KEY");
            var privateKey = Environment.GetEnvironmentVariable("ASPOSE_EMAIL_METERED_PRIVATE_KEY");

            if (string.IsNullOrEmpty(publicKey) || string.IsNullOrEmpty(privateKey))
            {
                Console.WriteLine("Metered licensing is not configured - skipping.");
                Console.WriteLine("Set ASPOSE_EMAIL_METERED_PUBLIC_KEY and ASPOSE_EMAIL_METERED_PRIVATE_KEY");
                Console.WriteLine("to run this example against your metered subscription.");
                return;
            }

            // SetMeteredKey validates the keys against the Aspose metering service, so it
            // needs both valid credentials and an internet connection.
            var metered = new Metered();
            metered.SetMeteredKey(publicKey, privateKey);

            Console.WriteLine("Metered license applied.");

            // Any API call from here on is counted against the metered subscription.
            var eml = MailMessage.Load(Data.Email/"Message.eml");
            Console.WriteLine($"Loaded message: {eml.Subject}");
        }
    }
}

// Demonstrates how to train a Bayesian spam filter using ham and spam email samples,
// then test messages to determine their spam probability.

using System;
using System.IO;
using Aspose.Email.AntiSpam;

namespace Aspose.Email.Examples.Email
{
    internal static class BayesianSpamAnalyzer
    {
        public static void Run()
        {
            var hamFolder = Data.Email/"hamFolder";
            var spamFolder = Data.Email/"Spam";
            var testFolder = Data.Email;

            // The trained database is generated output, so it belongs in Out rather than
            // in the input folder the examples read their sample messages from.
            var dataBaseFile = Data.Out/"SpamFilterDatabase.txt";

            TeachAndCreateDatabase(hamFolder, spamFolder, dataBaseFile);
            Console.WriteLine($"Trained filter saved to {dataBaseFile}\n");

            var analyzer = new SpamAnalyzer(dataBaseFile);
            var tested = 0;

            foreach (var file in Directory.GetFiles(testFolder, "*.eml"))
            {
                var msg = MailMessage.Load(file);
                Console.WriteLine($"{msg.Subject} -> {Describe(analyzer.Test(msg))}");
                tested++;
            }

            Console.WriteLine($"\nTested {tested} message(s).");
        }

        private static void TeachAndCreateDatabase(string hamFolder, string spamFolder, string dataBaseFile)
        {
            var analyzer = new SpamAnalyzer();

            foreach (var file in Directory.GetFiles(hamFolder, "*.eml"))
            {
                try { analyzer.TrainFilter(MailMessage.Load(file), false); }
                catch { continue; }
            }

            foreach (var file in Directory.GetFiles(spamFolder, "*.eml"))
            {
                try { analyzer.TrainFilter(MailMessage.Load(file), true); }
                catch { continue; }
            }

            analyzer.SaveDatabase(dataBaseFile);
        }

        // Test() returns the probability that the message is spam.
        private static string Describe(double probability)
        {
            if (probability < 0.05) return $"ham ({probability:P1} spam probability)";
            if (probability > 0.95) return $"spam ({probability:P1} spam probability)";
            return $"maybe spam ({probability:P1} spam probability)";
        }
    }
}

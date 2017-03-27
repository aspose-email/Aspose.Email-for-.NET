using System;
using System.IO;
using Aspose.Email.AntiSpam;
using Aspose.Email.Mime;


/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/

namespace Aspose.Email.Examples.CSharp.Email
{
    class BayesianSpamAnalyzer
    {
        public static void Run()
        {
            // ExStart:BayesianSpamAnalyzer
            string hamFolder = RunExamples.GetDataDir_Email() + "/hamFolder";
            string spamFolder = RunExamples.GetDataDir_Email() + "/Spam";
            string testFolder = RunExamples.GetDataDir_Email();
            string dataBaseFile = RunExamples.GetDataDir_Email() + "SpamFilterDatabase.txt";

            TeachAndCreateDatabase(hamFolder, spamFolder, dataBaseFile);
            string[] testFiles = Directory.GetFiles(testFolder, "*.eml");
            SpamAnalyzer analyzer = new SpamAnalyzer(dataBaseFile);
            foreach (string file in testFiles)
            {
                MailMessage msg = MailMessage.Load(file);
                Console.WriteLine(msg.Subject);
                double probability = analyzer.Test(msg);
                PrintResult(probability);
            }
            // ExEnd:BayesianSpamAnalyzer
        }

        private static void TeachAndCreateDatabase(string hamFolder, string spamFolder, string dataBaseFile)
        {
            string[] hamFiles = Directory.GetFiles(hamFolder, "*.eml");
            string[] spamFiles = Directory.GetFiles(spamFolder, "*.eml");

            SpamAnalyzer analyzer = new SpamAnalyzer();

            for (int i = 0; i < hamFiles.Length; i++)
            {
                MailMessage hamMailMessage;
                try
                {
                    hamMailMessage = MailMessage.Load(hamFiles[i]);
                }
                catch (Exception)
                {
                    continue;
                }

                Console.WriteLine(i);
                analyzer.TrainFilter(hamMailMessage, false);
            }

            for (int i = 0; i < spamFiles.Length; i++)
            {
                MailMessage spamMailMessage;
                try
                {
                    spamMailMessage = MailMessage.Load(hamFiles[i]);
                }
                catch (Exception)
                {
                    continue;
                }

                Console.WriteLine(i);
                analyzer.TrainFilter(spamMailMessage, true);
            }

            analyzer.SaveDatabase(dataBaseFile);
        }

        private static void PrintResult(double probability)
        {
            if (probability < 0.05)
                Console.WriteLine("This is ham)");
            else if (probability > 0.95)
                Console.WriteLine("This is spam)");
            else
                Console.WriteLine("Maybe spam)");
        }
    }
}

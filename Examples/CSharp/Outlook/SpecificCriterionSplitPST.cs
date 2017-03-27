using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using Aspose.Email.Storage.Pst;
using Aspose.Email.Tools.Search;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/

namespace Aspose.Email.Examples.CSharp.Email.Outlook
{
    class SpecificCriterionSplitPST
    {
        public static void Run()
        {
            // The path to the File directory.
            // ExStart:SpecificCriterionSplitPST
            string dataDir = RunExamples.GetDataDir_Outlook();
            IList<MailQuery> criteria = new List<MailQuery>();
            PersonalStorageQueryBuilder pstQueryBuilder = new PersonalStorageQueryBuilder();
            pstQueryBuilder.SentDate.Since(new DateTime(2005, 04, 01));
            pstQueryBuilder.SentDate.Before(new DateTime(2005, 04, 07));
            criteria.Add(pstQueryBuilder.GetQuery());
            pstQueryBuilder = new PersonalStorageQueryBuilder();
            pstQueryBuilder.SentDate.Since(new DateTime(2005, 04, 07));
            pstQueryBuilder.SentDate.Before(new DateTime(2005, 04, 13));
            criteria.Add(pstQueryBuilder.GetQuery());

            if (Directory.GetFiles(dataDir + "pathToPst", "*.pst").Length == 0)
            {
                
            }
            else
            {
                string[] files = Directory.GetFiles(dataDir + "pathToPst");

                foreach (string file in files)
                {
                    if(file.Contains(".pst"))
                    File.Delete(file);
                }
            }

            using (PersonalStorage personalStorage = PersonalStorage.FromFile(dataDir + "PersonalStorage_New.pst"))
            {
                personalStorage.SplitInto(criteria, dataDir + "pathToPst");
            }
            // ExEnd:SpecificCriterionSplitPST
        }
    }
}
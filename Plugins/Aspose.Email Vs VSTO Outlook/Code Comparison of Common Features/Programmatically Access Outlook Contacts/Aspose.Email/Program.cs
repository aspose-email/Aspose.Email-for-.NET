using Aspose.Email.Formats.VCard;
using Aspose.Email.Outlook;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Aspose.Email
{
    class Program
    {
        static void Main(string[] args)
        {
            var vcfTest = VCardContact.Load(@"E:\Aspose\Aspose Vs VSTO\Aspose.Emails Vs VSTO Outlook v 1.1\Sample Files\Jon.vcf");

            MapiContact contact = MapiContact.FromVCard(@"E:\Aspose\Aspose Vs VSTO\Aspose.Emails Vs VSTO Outlook v 1.1\Sample Files\Jon.vcf");

            Console.WriteLine(contact.NameInfo.DisplayName);
        }
    }
}

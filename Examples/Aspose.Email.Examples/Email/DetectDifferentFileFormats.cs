// Demonstrates how to detect the file format of an email message file
// using FileFormatUtil.

using System;
using Aspose.Email.Tools;

namespace Aspose.Email.Examples.Email
{
    internal static class DetectDifferentFileFormats
    {
        public static void Run()
        {
            var info = FileFormatUtil.DetectFileFormat(Data.Email/"Message.msg");
            Console.WriteLine($"The message format is: {info.FileFormatType}");
        }
    }
}

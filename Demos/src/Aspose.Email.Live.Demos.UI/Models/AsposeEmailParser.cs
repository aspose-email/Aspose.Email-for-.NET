using Aspose.Email.Live.Demos.UI.FileProcessing;
using Aspose.Email.Live.Demos.UI.LibraryHelpers;
using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Web.Http;
using Aspose.Email.Live.Demos.UI.Models.Common;

namespace Aspose.Email.Live.Demos.UI.Models
{
	///<Summary>
	/// AsposeEmailParser class to parse email file
	///</Summary>
	public class AsposeEmailParser : EmailBase
	{
		///<Summary>
		/// initialize AsposeEmailParserController class 
		///</Summary>
		public AsposeEmailParser()
        {
			Aspose.Email.Live.Demos.UI.Models.License.SetAsposeEmailLicense();
        }

		
		public Response Parse(InputFiles inputFiles)
		{
			(string folderName, string[] fileNames) pair = (null, null);

			pair.fileNames = inputFiles.FileName;
			pair.folderName = inputFiles.FolderName;

			try
			{
				
				return  Parse(Path.GetFileName(pair.fileNames[0]), pair.folderName);
			}
			catch (Exception ex)
			{
				Console.WriteLine(ex.Message);
				return new Response() { Status = "200", Text = "Error on processing file" };
			}
		}

		///<Summary>
		/// Parse method to parse email file
		///</Summary>
		public Response Parse(string inputFileName, string inputFolderName)
        {
            var processor = new CustomSingleOrZipFileProcessor()
            {
               
                CustomFileProcessMethod = ProcessEmailFileToResponce
            };

            return processor.Process(inputFolderName, inputFileName);
        }

        private void ProcessEmailFileToResponce(string inputFilePath, string outputFolderPath)
        {
            var mail = MapiHelper.GetMapiMessageFromFile(inputFilePath);

            if (mail == null)
                throw new Exception("Invalid file format");

            if (mail.Body != null)
                System.IO.File.AppendAllText(Path.Combine(outputFolderPath, "body.txt"), mail.Body);
            if (mail.BodyRtf != null)
                System.IO.File.AppendAllText(Path.Combine(outputFolderPath, "body.rtf"), mail.BodyRtf);
            if (mail.BodyHtml != null)
                System.IO.File.AppendAllText(Path.Combine(outputFolderPath, "body.html"), mail.BodyHtml);

            if (mail.Attachments != null)
                foreach (var attachment in mail.Attachments)
					attachment.Save(Path.Combine(outputFolderPath, Path.GetFileName(attachment.LongFileName)));
        }
    }
}

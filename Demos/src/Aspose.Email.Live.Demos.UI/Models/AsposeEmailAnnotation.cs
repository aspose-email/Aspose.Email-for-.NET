using Aspose.Email.Live.Demos.UI.FileProcessing;
using Aspose.Email.Live.Demos.UI.LibraryHelpers;
using Aspose.Email.Live.Demos.UI.Models;
using Aspose.Email;
using Aspose.Email.Mapi;
using System;
using System.IO;
using System.Threading.Tasks;
using System.Web.Http;
using Aspose.Email.Live.Demos.UI.Models.Common;

namespace Aspose.Email.Live.Demos.UI.Models
{
    public class AsposeEmailAnnotation : EmailBase
	{
		///<Summary>
		/// Remove annotations method
		///</Summary>
	
		public Response Remove(InputFiles inputFiles)
		{
			var folderName = inputFiles.FolderName;
			var fileNames = inputFiles.FileName;

			var processor = new CustomSingleOrZipFileProcessor()
			{
				
				CustomFilesProcessMethod = (string[] inputFilePaths, string outputFolderPath) =>
				{
					Aspose.Email.Live.Demos.UI.Models.License.SetAsposeEmailLicense();

					for (int i = 0; i < inputFilePaths.Length; i++)
					{
						var inputFilePath = inputFilePaths[i];

						var mail = MapiHelper.GetMapiMessageFromFile(inputFilePath);

						if (mail == null)
							throw new Exception("Invalid file format");

						var options = FollowUpManager.GetOptions(mail);

						var flag = options.FlagRequest;

						if (flag != null)
						{
							System.IO.File.WriteAllText(Path.Combine(outputFolderPath, "comment.txt"), $"{options.StartDate}{Environment.NewLine}{flag}");
							FollowUpManager.ClearFlag(mail);
						}

						mail.Save(Path.Combine(outputFolderPath, Path.GetFileNameWithoutExtension(inputFilePath) + ".msg"), MsgSaveOptions.DefaultMsgUnicode);
					}
				}
			};

			return processor.Process(folderName, fileNames);
		}

			
    }
}

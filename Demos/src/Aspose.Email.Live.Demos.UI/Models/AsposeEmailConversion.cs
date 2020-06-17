using System;
using System.IO;
using System.Web.Http;
using System.Threading.Tasks;
using System.Diagnostics;
using System.Linq;
using Aspose.Email.Live.Demos.UI.Controllers;
using Aspose.Email.Live.Demos.UI.FileProcessing;
using Aspose.Email.Live.Demos.UI.Models.Common;
using Aspose.Email.Live.Demos.UI.Services;
using Aspose.Email.Live.Demos.UI.Services.Email;

namespace Aspose.Email.Live.Demos.UI.Models
{
	///<Summary>
	/// AsposeCellsConversion class to convert words files to different formats
	///</Summary>
	public class AsposeEmailConversion : EmailBase
	{

		public Response Convert(InputFiles inputFiles, string outputType)
		{
			(string folderName, string[] fileNames) pair = (null, null);

			pair.fileNames = inputFiles.FileName;
			pair.folderName = inputFiles.FolderName;
			try
			{
				

				var processor = new CustomSingleOrZipFileProcessor()
				{
					
					CustomFilesProcessMethod = (files, outputFolder) =>
					{
						var handler = new FolderOutputHandler(outputFolder);

						using (var service = new EmailService())
						{
							for (int i = 0; i < files.Length; i++)
							{
								var filePath = files[i];
								var fileName = Path.GetFileName(filePath);

								using (var input = System.IO.File.OpenRead(filePath))
									service.Convert(input, fileName, handler, outputType);
							}
						}
					}
				};

				return processor.Process(pair.folderName, pair.fileNames);
			}
			catch (Exception ex)
			{
				Console.WriteLine(ex.Message);
				return new Response() { Status = "200", Text = "Error on processing file" };
			}
		}


	}
}

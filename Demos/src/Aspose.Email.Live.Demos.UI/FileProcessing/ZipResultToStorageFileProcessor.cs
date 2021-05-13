using Aspose.Email.Live.Demos.UI.Models;
using Aspose.Email.Live.Demos.UI.Services;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Threading.Tasks;

namespace Aspose.Email.Live.Demos.UI.FileProcessing
{
    ///<Summary>
    /// SingleOrZipFileProcessor class to check either to create zip file or not
    ///</Summary>
    public abstract class ZipResultToStorageFileProcessor : LoggableFileProcessor
    {
		public IStorageService StorageService { get; }

		protected ZipResultToStorageFileProcessor(IStorageService storageService, IConfiguration configuration, ILogger logger) 
			: base(logger, configuration)
        {
			StorageService = storageService;
		}

        protected override async Task ProcessFilesToResponce(Response resp, IDictionary<string, byte[]> files)
		{
			var handler = new InMemoryOutputHandler();

			ProcessFiles(files, handler);
			
			var folder = Guid.NewGuid().ToString();

			// Dont create Zip if there is only one file in output
			if (handler.Resources.Count == 1)
			{
				var pair = handler.Resources.First();

				var fileName = Path.GetFileName(pair.Key);
				var stream = pair.Value;
				stream.Position = 0;

				await StorageService.SaveFile(stream, folder, fileName);

				resp.FolderName = folder;
				resp.FileName = fileName;
			}
			else
			{
				var zipFileName = Guid.NewGuid().ToString() + ".zip";

				using (var stream = new MemoryStream())
				{
					using (var archive = new ZipArchive(stream, ZipArchiveMode.Create, true))
					{
                        foreach (var pair in handler.Resources)
                        {
							var file = archive.CreateEntry(pair.Key);
							using (var entryStream = file.Open())
							{
								pair.Value.Position = 0;
								await pair.Value.CopyToAsync(entryStream);
							}
						}
					}

					stream.Position = 0;
					await StorageService.SaveFile(stream, folder, zipFileName);
				}

				resp.FolderName = folder;
				resp.FileName = zipFileName;
			}
		}

		///<Summary>
		/// ProcessFileToPath method to process files to path
		///</Summary>
		/// <param name="inputFilePaths"></param>
		/// <param name="outDirectoryPath"></param>
		protected abstract void ProcessFiles(IDictionary<string, byte[]> files, IOutputHandler handler);
	}
}

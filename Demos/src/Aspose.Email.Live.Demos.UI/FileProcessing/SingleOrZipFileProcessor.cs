using Aspose.Email.Live.Demos.UI.Models;
using Aspose.Email.Live.Demos.UI.Config;
using System;
using System.IO;
using System.IO.Compression;
using System.Linq;

namespace Aspose.Email.Live.Demos.UI.FileProcessing
{
	///<Summary>
	/// SingleOrZipFileProcessor class to check either to create zip file or not
	///</Summary>
	public abstract class SingleOrZipFileProcessor : FileProcessor
    {
		///<Summary>
		/// ProcessFileToResponce method 
		///</Summary>
		/// <param name="resp"></param>
		/// <param name="inputFolderName"></param>
		/// <param name="inputFileName"></param>
		/// <param name="outputFileName"></param>

		protected override void ProcessFileToResponce(Response resp, string inputFolderName, string inputFileName, string outputFileName = null)
		{
			var inputFolderPath = Path.Combine(Config.Configuration.WorkingDirectory, inputFolderName);
			var inputFilePath = Path.Combine(inputFolderPath, inputFileName);

			var outputFolderName = Guid.NewGuid().ToString();
			var outputFolderPath = Path.Combine(Config.Configuration.OutputDirectory, outputFolderName);

			Directory.CreateDirectory(outputFolderPath);

			ProcessFileToPath(inputFilePath, outputFolderPath);

			var files = Directory.GetFiles(outputFolderPath);

			// Dont create Zip if there is only one file in output
			if (files.Length == 1)
			{
				outputFileName = outputFileName ?? Path.GetFileName(files[0]);

				resp.FileName = outputFileName;
				resp.FolderName = outputFolderName;
			}
			else
			{
				var outputZipFolderName = Guid.NewGuid().ToString();
				var outputZipFolderPath = Path.Combine(Config.Configuration.OutputDirectory, outputZipFolderName);
				Directory.CreateDirectory(outputZipFolderPath);

				outputFileName = (outputFileName ?? Guid.NewGuid().ToString()) + ".zip";

				var outputZipFilePath = Path.Combine(outputZipFolderPath, outputFileName);

				ZipFile.CreateFromDirectory(outputFolderPath, outputZipFilePath);
				Directory.Delete(outputFolderPath, true);

				resp.FileName = outputFileName;
				resp.FolderName = outputZipFolderName;
			}
		}

		protected override void ProcessFileToResponce(Response resp, string inputFolderName, string[] inputFileNames)
		{
			var inputFolderPath = Path.Combine(Config.Configuration.WorkingDirectory, inputFolderName);
			var inputFilePaths = inputFileNames.Select(x => Path.Combine(inputFolderPath, x)).ToArray();

			var outputFolderName = Guid.NewGuid().ToString();
			var outputFolderPath = Path.Combine(Config.Configuration.OutputDirectory, outputFolderName);

			Directory.CreateDirectory(outputFolderPath);

			ProcessFilesToPath(inputFilePaths, outputFolderPath);

			var files = Directory.GetFiles(outputFolderPath);

			// Dont create Zip if there is only one file in output
			if (files.Length == 1)
			{
				var outputFileName = Path.GetFileName(files[0]);

				resp.FileName = outputFileName;
				resp.FolderName = outputFolderName;
			}
			else
			{
				var outputZipFolderName = Guid.NewGuid().ToString();
				var outputZipFolderPath = Path.Combine(Config.Configuration.OutputDirectory, outputZipFolderName);
				Directory.CreateDirectory(outputZipFolderPath);

				var outputFileName = Guid.NewGuid().ToString() + ".zip";

				var outputZipFilePath = Path.Combine(outputZipFolderPath, outputFileName);

				ZipFile.CreateFromDirectory(outputFolderPath, outputZipFilePath);
				Directory.Delete(outputFolderPath, true);

				resp.FileName = outputFileName;
				resp.FolderName = outputZipFolderName;
			}
		}

		///<Summary>
		/// ProcessFileToPath method to process file to path
		///</Summary>
		/// <param name="inputFilePath"></param>
		/// <param name="outDirectoryPath"></param>
		protected abstract void ProcessFileToPath(string inputFilePath, string outDirectoryPath);

		///<Summary>
		/// ProcessFileToPath method to process files to path
		///</Summary>
		/// <param name="inputFilePaths"></param>
		/// <param name="outDirectoryPath"></param>
		protected abstract void ProcessFilesToPath(string[] inputFilePaths, string outDirectoryPath);
	}
}

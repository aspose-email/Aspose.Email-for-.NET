namespace Aspose.Email.Live.Demos.UI.FileProcessing
{
	///<Summary>
	/// CustomSingleOrZipFileProcessor class 
	///</Summary>
	public class CustomSingleOrZipFileProcessor : SingleOrZipFileProcessor
    {
		///<Summary>
		/// ProcessFileDelegate delegate
		///</Summary>
		public delegate void ProcessFileDelegate(string inputFilePath, string outputFolderPath);

		///<Summary>
		/// ProcessFilesDelegate delegate
		///</Summary>
		public delegate void ProcessFilesDelegate(string[] inputFilePath, string outputFolderPath);

		///<Summary>
		/// CustomProcessMethod method
		///</Summary>
		public ProcessFileDelegate CustomFileProcessMethod { get; set; }

		///<Summary>
		/// CustomProcessMethod method
		///</Summary>
		public ProcessFilesDelegate CustomFilesProcessMethod { get; set; }

		///<Summary>
		/// ProcessFileToPath method
		///</Summary>
		protected override void ProcessFileToPath(string inputFilePath, string outputFolderPath)
		{
			var action = CustomFileProcessMethod;

			action?.Invoke(inputFilePath, outputFolderPath);
		}

		///<Summary>
		/// ProcessFilesToPath method
		///</Summary>
		protected override void ProcessFilesToPath(string[] inputFilePaths, string outputFolderPath)
		{
			var action = CustomFilesProcessMethod;

			action?.Invoke(inputFilePaths, outputFolderPath);
		}


	}
}

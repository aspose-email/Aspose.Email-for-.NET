
using Aspose.Email.Live.Demos.UI.Services;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using System.Collections.Generic;

namespace Aspose.Email.Live.Demos.UI.FileProcessing
{
	///<Summary>
	/// CustomSingleOrZipFileProcessor class 
	///</Summary>
	public class CustomSingleOrZipFileProcessor : ZipResultToStorageFileProcessor
    {
        public CustomSingleOrZipFileProcessor(IStorageService storageService, IConfiguration configuration, ILogger logger) 
			: base(storageService, configuration, logger)
        {
        }

        ///<Summary>
        /// ProcessFilesDelegate delegate
        ///</Summary>
        public delegate void ProcessFilesDelegate(IDictionary<string, byte[]> files, IOutputHandler handler);

		///<Summary>
		/// CustomProcessMethod method
		///</Summary>
		public ProcessFilesDelegate CustomFilesProcessMethod { get; set; }

		///<Summary>
		/// ProcessFilesToPath method
		///</Summary>
		protected override void ProcessFiles(IDictionary<string, byte[]> files, IOutputHandler handler)
		{
			CustomFilesProcessMethod?.Invoke(files, handler);
		}
	}
}

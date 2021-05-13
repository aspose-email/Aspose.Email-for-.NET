using Aspose.Email.Live.Demos.UI.Models;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace Aspose.Email.Live.Demos.UI.FileProcessing
{
    ///<Summary>
    /// Base class for implementing file processing logic
    ///</Summary>
    public abstract class FileProcessor
    {
        ///<Summary>
        /// Process file in input and return Response 
        ///</Summary>
        ///<param name="inputFolderName">Input folder with files</param>
        ///<param name="inputFileNames">Short file names</param>
        public virtual async Task<Response> Process(IDictionary<string, byte[]> files)
		{
			var resp = new Response()
			{
				StatusCode = 200,
				Status = "OK"
			};

			await ProcessFilesToResponce(resp, files);

			return resp;
		}
		///<Summary>
		/// Process files in input and modify Response 
		///</Summary>
		///<param name="inputFolderName"></param>
		///<param name="inputFileNames">Short file names</param>
		///<param name="resp">Responce to modify</param>
		protected abstract Task ProcessFilesToResponce(Response resp, IDictionary<string, byte[]> files);
	}
}

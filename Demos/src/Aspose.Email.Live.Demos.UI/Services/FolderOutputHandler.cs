using System.IO;

namespace Aspose.Email.Live.Demos.UI.Services
{
	public class FolderOutputHandler : IOutputHandler
	{
        readonly string _outputFolder;

        public FolderOutputHandler(string outputFolder)
        {
            _outputFolder = outputFolder;
        }

        public Stream CreateOutputStream(string name)
		{
			return File.Create(Path.Combine(_outputFolder, name));
		}

		public void RemoveResource(string name)
		{
			var path = Path.Combine(_outputFolder, name);
			if (File.Exists(path))
				File.Delete(path);
		}
	}
}

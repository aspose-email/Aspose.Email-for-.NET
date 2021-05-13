using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace Aspose.Email.Live.Demos.UI.Services.Api.Storage
{
    public class FileStorageService : IStorageService
    {
        string Path(params string[] parts) => System.IO.Path.Combine(parts);

        public Task<bool> IsFileExists(string folder, string name)
        {
            return File.Exists(Path("Temp",folder, name)).ToTaskResult();
        }

        public Task<Stream> ReadFile(string folder, string name)
        {
            return File.OpenRead(Path("Temp", folder, name)).ToTaskResult<Stream>();
        }

        public Task<IEnumerable<(string name, Task<Stream> data)>> ReadFilesFromFolder(string folder)
        {
            Directory.CreateDirectory(Path("Temp", folder));
            return Directory.EnumerateFiles(Path("Temp", folder)).Select(x => (System.IO.Path.GetFileName(x), new Task<Stream>(() => File.OpenRead(x)))).ToTaskResult();
        }

        public Task SaveFile(string content, string folder, string name)
        {
            Directory.CreateDirectory(Path("Temp", folder));
            return File.WriteAllTextAsync(Path("Temp", folder, name), content);
        }

        public Task SaveFile(byte[] content, string folder, string name)
        {
            Directory.CreateDirectory(Path("Temp", folder));
            return File.WriteAllBytesAsync(Path("Temp", folder, name), content);
        }

        public Task SaveFile(Stream content, string folder, string name)
        {
            Directory.CreateDirectory(Path("Temp", folder));
            var memoryStream = content.CopyToMemoryStream();
            return File.WriteAllBytesAsync(Path("Temp", folder, name), memoryStream.ToArray());
        }
    }
}

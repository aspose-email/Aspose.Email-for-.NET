using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;

namespace Aspose.Email.Live.Demos.UI.Services
{
    public interface IStorageService
    {
        Task SaveFile(string content, string folder, string name);
        Task SaveFile(byte[] content, string folder, string name);
        Task SaveFile(Stream content, string folder, string name);
        Task<Stream> ReadFile(string folder, string name);
        Task<IEnumerable<(string name, Task<Stream> data)>> ReadFilesFromFolder(string folder);
        Task<bool> IsFileExists(string folder, string name);
    }
}

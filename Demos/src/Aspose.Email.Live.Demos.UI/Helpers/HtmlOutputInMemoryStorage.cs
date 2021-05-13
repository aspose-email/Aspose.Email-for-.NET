using Aspose.Html.IO;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace Aspose.Email.Live.Demos.UI.Helpers
{
    public class HtmlOutputInMemoryStorage : IOutputStorage
    {
        Dictionary<string, MemoryStream> _files = new Dictionary<string, MemoryStream>();

        public string ReadResourceAsString(int index)
        {
            var stream = _files.Values.ElementAt(index);
            stream.Position = 0;

            using (var reader = new StreamReader(stream))
            {
                var content = reader.ReadToEnd();
                content = content.Replace("<head></head><body></body>", "");
                return content;
            }
        }

        public int ResourcesCount => _files.Count;

        public OutputStream CreateStream(OutputStreamContext context)
        {
            return new OutputStream(_files[context.Uri] = new MemoryStream(), context.Uri);
        }

        public void ReleaseStream(OutputStream stream)
        {
        }
    }
}

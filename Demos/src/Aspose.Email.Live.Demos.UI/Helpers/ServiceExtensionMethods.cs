using Aspose.Email.Live.Demos.UI.Services;
using Aspose.Email.Mapi;
using System.IO;
using System.Threading.Tasks;

namespace Aspose.Email.Live.Demos.UI.Helpers
{
    public static class ServiceExtensionMethods
    {
        public static async Task<MapiMessage> ReadMapiMessage(this IStorageService service, string uid, string name)
        {
            using (var stream = await service.ReadFile(uid, name))
                return MapiHelper.GetMapiMessageFromStream(stream);
        }

        public static Task SaveMessage(this IStorageService service, MapiMessage message, string uid, string name, SaveOptions options)
        {
            var stream = new MemoryStream();
            message.Save(stream, options);
            stream.Position = 0;
            return service.SaveFile(stream, uid, name);
        }
    }
}

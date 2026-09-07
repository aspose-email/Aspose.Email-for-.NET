// Demonstrates the streaming counterparts of GetSubFolders and GetContents, plus the
// two kinds of item a PST folder holds.
//
// EnumerateFolders yields folders one at a time instead of building a collection, and
// FolderKind separates ordinary folders from search folders. MessageKind does the same
// for items: Normal is what the user sees, FolderAssociatedInformation is the hidden
// configuration Outlook keeps beside it (views, rules, forms).

using System;
using System.Linq;
using Aspose.Email.Storage.Pst;

namespace Aspose.Email.Examples.PST
{
    internal static class EnumerateFoldersAndMessageKinds
    {
        public static void Run()
        {
            using (var pst = PersonalStorage.FromFile(Data.Mapi/"Sub.pst", false))
            {
                var all = pst.RootFolder.EnumerateFolders().Count();
                var normal = pst.RootFolder.EnumerateFolders(FolderKind.Normal).Count();
                var search = pst.RootFolder.EnumerateFolders(FolderKind.Search).Count();

                Console.WriteLine($"Folders under the root: {all} ({normal} normal, {search} search)");
                Console.WriteLine();

                foreach (var folder in pst.RootFolder.EnumerateFolders())
                {
                    var visible = folder.GetContents(MessageKind.Normal).Count;
                    var associated = folder.GetContents(MessageKind.FolderAssociatedInformation).Count;

                    Console.WriteLine($"{folder.RetrieveFullPath()}");
                    Console.WriteLine($"   container class: {folder.ContainerClass}");
                    Console.WriteLine($"   normal items:    {visible}");
                    Console.WriteLine($"   associated:      {associated}");
                }
            }
        }
    }
}

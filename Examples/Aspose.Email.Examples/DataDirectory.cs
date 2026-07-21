using System;
using System.IO;

namespace Aspose.Email.Examples
{

    internal static class Data
    {
        private static readonly char Sep = Path.DirectorySeparatorChar;

        // Each property maps to a folder under Data\, named after the example
        // category that uses it.

        internal static SubDirectory Mbox
        {
            get
            {
               return new SubDirectory(Path.GetFullPath(GetPath() + $@"MBOX{Sep}"));
            }
        }

        internal static SubDirectory Email
        {
            get
            {
                return new SubDirectory(Path.GetFullPath(GetPath() + $@"Email{Sep}"));
            }
        }

        internal static SubDirectory Ews
        {
            get
            {
                return new SubDirectory(Path.GetFullPath(GetPath() + $@"EWS{Sep}"));
            }
        }

        internal static SubDirectory Mapi
        {
            get
            {
                return new SubDirectory(Path.GetFullPath(GetPath() + $@"MAPI{Sep}"));
            }
        }

        internal static SubDirectory Pop3
        {
            get
            {
                return new SubDirectory(Path.GetFullPath(GetPath() + $@"POP3{Sep}"));
            }
        }

        internal static SubDirectory Imap
        {
            get
            {
                return new SubDirectory(Path.GetFullPath(GetPath() + $@"IMAP{Sep}"));
            }
        }

        internal static SubDirectory Smtp
        {
            get
            {
                return new SubDirectory(Path.GetFullPath(GetPath() + $@"SMTP{Sep}"));
            }
        }

        internal static SubDirectory Gmail
        {
            get
            {
                return new SubDirectory(Path.GetFullPath(GetPath() + $@"Gmail{Sep}"));
            }
        }

        internal static SubDirectory Out
        {
            get
            {
                return new SubDirectory(Path.GetFullPath(GetOutPath()));
            }
        }

        // Returns an Out\<name>\ subdirectory, creating it if it does not exist yet.
        // Use this instead of Data.Out/"name" whenever an example writes into a
        // subfolder: the Out directory is emptied on every run, so the subfolder
        // has to be recreated before anything can be written to it.
        internal static SubDirectory OutSub(string name)
        {
            var path = Path.Combine(GetOutPath(), name);
            Directory.CreateDirectory(path);
            return new SubDirectory(path + Sep);
        }

        internal static void ClearOut()
        {
            var outPath = GetOutPath();

            Directory.CreateDirectory(outPath);

            foreach (var file in Directory.EnumerateFiles(outPath))
            {
                File.Delete(file);
            }
            foreach (var dir in Directory.EnumerateDirectories(outPath))
            {
                // Recursive: examples create non-empty subfolders (Contacts,
                // Attachments, ...) and a non-recursive delete would throw.
                Directory.Delete(dir, true);
            }
        }

        // The solution root. Also used to locate the Aspose.Email license file.
        internal static string RootPath => GetRootPath();

        private static string GetPath()
        {
            return Path.Combine(RequireRootPath(), $@"Data{Sep}");
        }

        private static string GetOutPath()
        {
            return Path.Combine(RequireRootPath(), $@"Out{Sep}");
        }

        private static string RequireRootPath()
        {
            return GetRootPath() ?? throw new InvalidOperationException(
                "Could not locate the solution root. Run the examples from inside the " +
                "solution directory - the root is the folder containing both a .sln " +
                "file and the Data folder.");
        }

        private static string GetRootPath()
        {
            // Walk up from the current working directory until we find the
            // solution root: a directory that contains a .sln file alongside
            // a Data/ folder.  This works whether the app is run from the
            // bin/Debug/net8.0/ output directory (exe double-click) or from
            // the project directory (dotnet run).
            var dir = new DirectoryInfo(Directory.GetCurrentDirectory());
            while (dir != null)
            {
                if (dir.GetFiles("*.sln").Length > 0 &&
                    Directory.Exists(Path.Combine(dir.FullName, "Data")))
                {
                    return dir.FullName;
                }
                dir = dir.Parent;
            }
            return null;
        }
    }

    internal class SubDirectory
    {
        private readonly string path;

        internal SubDirectory(string path)
        {
            this.path = path;
        }

        public override string ToString()
        {
            return this.path;
        }

        public static implicit operator string(SubDirectory subDirectory) => subDirectory.path;

        internal string GetFileName(string fileName) => Path.Combine(path, fileName);

        public static string operator /(SubDirectory subDirectory, string fileName)
        {
            return Path.Combine(subDirectory.path, fileName);
        }
    }
}
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;

namespace Aspose.Email.Examples
{
    internal static class ExampleRunner
    {
        private const int PageSize = 20;

        // Sentinel returned by ShowExampleList to signal the user chose to quit.
        private static readonly ExampleEntry QuitSentinel = new ExampleEntry();

        internal static void Run(string[] args)
        {
            if (args.Length > 0)
            {
                // --bare: suppress the decorative header/footer so the GUI shell
                // can apply its own formatting.  Exit code 1 signals an example error.
                bool bare = args.Any(a => string.Equals(a, "--bare", StringComparison.Ordinal));
                string name = string.Join(" ", args.Where(a => a != "--bare"));
                RunByName(name, bare);
                return;
            }

            if (Console.IsInputRedirected || Console.IsOutputRedirected)
            {
                // Non-interactive context (CI, piped output). Print the example list and exit.
                PrintList(DiscoverExamples());
                return;
            }

            RunInteractive();
        }

        private static void PrintList(List<ExampleEntry> examples)
        {
            Console.WriteLine("Aspose.Email for .NET Examples (" + examples.Count + " available)");
            Console.WriteLine();
            string currentCategory = null;
            foreach (var ex in examples)
            {
                if (ex.Category != currentCategory)
                {
                    Console.WriteLine("[" + ex.Category + "]");
                    currentCategory = ex.Category;
                }
                Console.WriteLine("  " + ex.Name);
            }
            Console.WriteLine();
            Console.WriteLine("Run with an example name as argument to execute it:");
            Console.WriteLine("  dotnet run -- CreateNewEmail");
        }

        // ------------------------------------------------------------------
        // Discovery
        // ------------------------------------------------------------------

        private static List<ExampleEntry> DiscoverExamples()
        {
            // Find every public static Run() method in the assembly.
            // This covers all example classes without requiring any registration.
            return Assembly.GetExecutingAssembly()
                .GetTypes()
                .Select(t => new
                {
                    Type = t,
                    Method = t.GetMethod("Run",
                        BindingFlags.Public | BindingFlags.Static,
                        null,
                        Type.EmptyTypes,
                        null)
                })
                .Where(x => x.Method != null)
                .Select(x => new ExampleEntry(x.Type, x.Method))
                .OrderBy(e => e.Category)
                .ThenBy(e => e.Name)
                .ToList();
        }

        // ------------------------------------------------------------------
        // CLI mode
        // ------------------------------------------------------------------

        private static void RunByName(string name, bool bare = false)
        {
            var examples = DiscoverExamples();
            var match = examples.FirstOrDefault(e =>
                string.Equals(e.Name, name, StringComparison.OrdinalIgnoreCase) ||
                string.Equals(e.FullName, name, StringComparison.OrdinalIgnoreCase));

            if (match == null)
            {
                Console.Error.WriteLine("No example found matching '" + name + "'.");

                // In bare mode the caller relies on the exit code, so an unknown
                // name has to fail rather than look like a successful run.
                if (bare) System.Environment.Exit(1);

                Console.WriteLine("Run without arguments to browse examples interactively.");
                return;
            }

            // clearScreen: false — CLI mode keeps the terminal's scroll history intact.
            bool ok = ExecuteExample(match, clearScreen: false, bare: bare);

            // In bare mode let the caller distinguish success from error via exit code.
            if (bare && !ok)
                System.Environment.Exit(1);
        }

        // ------------------------------------------------------------------
        // Interactive mode
        // ------------------------------------------------------------------

        private static void RunInteractive()
        {
            var allExamples = DiscoverExamples();

            while (true)
            {
                var filtered = ShowFilterPrompt(allExamples);
                if (filtered == null)   // user typed 'q' at the filter prompt
                    return;

                if (filtered.Count == 0)
                {
                    Console.WriteLine("No examples matched. Press any key to try again.");
                    Console.ReadKey();
                    continue;
                }

                var chosen = ShowExampleList(filtered);
                if (chosen == null)             // user pressed 'b' (back)
                    continue;
                if (ReferenceEquals(chosen, QuitSentinel))
                    return;

                ExecuteExample(chosen);
                Console.WriteLine();
                Console.Write("Press any key to continue...");
                Console.ReadKey();
            }
        }

        // Returns null when the user quits (types 'q').
        // Returns the filtered list (possibly the full list) otherwise.
        private static List<ExampleEntry> ShowFilterPrompt(List<ExampleEntry> all)
        {
            TryClear();
            WriteHeader("Aspose.Email for .NET  -  Example Runner");
            Console.WriteLine("  " + all.Count + " examples available\n");

            // Category summary line
            var byCategory = all
                .GroupBy(e => e.Category)
                .OrderBy(g => g.Key)
                .ToList();

            foreach (var group in byCategory)
                Console.WriteLine("  " + group.Key + " (" + group.Count() + ")");

            Console.WriteLine();
            Console.Write("Filter by keyword or category (Enter = show all, 'q' = quit): ");
            var input = (Console.ReadLine() ?? "").Trim();

            if (string.Equals(input, "q", StringComparison.OrdinalIgnoreCase))
                return null;

            if (string.IsNullOrEmpty(input))
                return all;

            return all
                .Where(e =>
                    ContainsCI(e.Name, input) ||
                    ContainsCI(e.Category, input))
                .ToList();
        }

        // Returns null  → go back to filter prompt
        // Returns QuitSentinel → exit the application
        // Returns an ExampleEntry → run that example
        private static ExampleEntry ShowExampleList(List<ExampleEntry> filtered)
        {
            int page = 0;
            int totalPages = (int)Math.Ceiling(filtered.Count / (double)PageSize);

            while (true)
            {
                TryClear();
                WriteHeader(
                    filtered.Count + " result(s)  [page " + (page + 1) + " / " + totalPages + "]");

                var pageItems = filtered.Skip(page * PageSize).Take(PageSize).ToList();
                int startIdx = page * PageSize;
                string currentCategory = null;

                for (int i = 0; i < pageItems.Count; i++)
                {
                    var ex = pageItems[i];
                    if (ex.Category != currentCategory)
                    {
                        Console.ForegroundColor = ConsoleColor.Cyan;
                        Console.WriteLine("\n  [" + ex.Category + "]");
                        Console.ResetColor();
                        currentCategory = ex.Category;
                    }

                    Console.WriteLine("  " + PadLeft(startIdx + i + 1, 4) + ".  " + ex.Name);
                }

                Console.WriteLine();

                if (totalPages > 1)
                    Console.WriteLine("  'n' next page   'p' prev page");

                Console.WriteLine("  'b' back to filter   'q' quit");
                Console.Write("\nEnter number to run: ");
                var input = (Console.ReadLine() ?? "").Trim();

                if (string.Equals(input, "q", StringComparison.OrdinalIgnoreCase))
                    return QuitSentinel;

                if (string.Equals(input, "b", StringComparison.OrdinalIgnoreCase))
                    return null;

                if (string.Equals(input, "n", StringComparison.OrdinalIgnoreCase) && page < totalPages - 1)
                {
                    page++;
                    continue;
                }

                if (string.Equals(input, "p", StringComparison.OrdinalIgnoreCase) && page > 0)
                {
                    page--;
                    continue;
                }

                if (int.TryParse(input, out int choice) && choice >= 1 && choice <= filtered.Count)
                    return filtered[choice - 1];

                Console.WriteLine("Invalid input. Press any key.");
                Console.ReadKey();
            }
        }

        // ------------------------------------------------------------------
        // Execution
        // ------------------------------------------------------------------

        // Returns true on success, false on error.
        private static bool ExecuteExample(ExampleEntry example, bool clearScreen = true, bool bare = false)
        {
            if (clearScreen) TryClear();

            if (!bare)
            {
                WriteHeader("Running: " + example.FullName);
                Console.WriteLine();
            }

            try
            {
                example.RunMethod.Invoke(null, null);

                if (!bare)
                {
                    Console.WriteLine();
                    Console.ForegroundColor = ConsoleColor.Green;
                    Console.WriteLine("Completed successfully.");
                    Console.ResetColor();
                }

                return true;
            }
            catch (TargetInvocationException tie)
            {
                var inner = tie.InnerException ?? tie;
                if (bare) Console.Error.WriteLine("Error: " + inner.Message);
                else PrintError(inner.Message);
                return false;
            }
            catch (Exception ex)
            {
                if (bare) Console.Error.WriteLine("Error: " + ex.Message);
                else PrintError(ex.Message);
                return false;
            }
        }

        private static void PrintError(string message)
        {
            Console.WriteLine();
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine("Error: " + message);
            Console.ResetColor();
        }

        // ------------------------------------------------------------------
        // Helpers
        // ------------------------------------------------------------------

        private static void TryClear()
        {
            try { Console.Clear(); }
            catch (IOException) { /* not a real console; skip clear */ }
        }

        private static void WriteHeader(string title)
        {
            Console.ForegroundColor = ConsoleColor.Yellow;
            Console.WriteLine(title);
            int width = SafeWindowWidth();
            Console.WriteLine(new string('-', Math.Min(title.Length + 2, width - 1)));
            Console.ResetColor();
        }

        private static int SafeWindowWidth()
        {
            try { return Console.WindowWidth; }
            catch { return 80; }
        }

        // .NET Framework 4.8 does not have string.Contains(string, StringComparison).
        private static bool ContainsCI(string source, string value) =>
            source.IndexOf(value, StringComparison.OrdinalIgnoreCase) >= 0;

        private static string PadLeft(int value, int width)
        {
            var s = value.ToString();
            return s.Length >= width ? s : new string(' ', width - s.Length) + s;
        }
    }

    // -----------------------------------------------------------------------
    // ExampleEntry
    // -----------------------------------------------------------------------

    internal sealed class ExampleEntry
    {
        internal string Category { get; }
        internal string Name { get; }
        internal string FullName { get { return Category + "/" + Name; } }
        internal MethodInfo RunMethod { get; }

        // Used only to create the QuitSentinel — all properties stay null/default.
        internal ExampleEntry() { }

        internal ExampleEntry(Type type, MethodInfo runMethod)
        {
            Category = DeriveCategory(type);
            Name = type.Name;
            RunMethod = runMethod;
        }

        // The category is simply the folder the example lives in: every example
        // namespace is "Aspose.Email.Examples.<Folder>", so the part after the
        // root namespace is the label.
        private static string DeriveCategory(Type type)
        {
            const string root = "Aspose.Email.Examples";

            var ns = type.Namespace ?? string.Empty;

            if (ns.StartsWith(root + ".", StringComparison.Ordinal))
                return ns.Substring(root.Length + 1);

            // Infrastructure types sit directly in the root namespace.
            return ns == root ? "General" : ns;
        }
    }
}

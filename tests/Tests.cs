using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using ExplorerMerge;

static class Tests
{
    static int checks;
    static void Check(bool condition, string label)
    {
        if (!condition) throw new Exception("FAIL: " + label);
        checks++;
        Console.WriteLine("PASS: " + label);
    }
    static Tab T(long window, long handle, string path)
    {
        return new Tab { Window = new IntPtr(window), Handle = new IntPtr(handle), Path = path };
    }
    static void PolicyTests()
    {
        var source = new List<Tab> { T(1, 10, "A"), T(1, 11, "B") };
        var copies = new List<Tab> { T(2, 20, "A"), T(2, 21, "B") };
        Check(MergePolicy.SameTabs(source, source.AsEnumerable().Reverse().ToList()), "Source tab ordering does not affect verification");
        Check(!MergePolicy.SameTabs(source, new List<Tab> { T(1, 10, "A") }), "Missing source tab prevents closing");
        Check(!MergePolicy.SameTabs(source, new List<Tab> { T(1, 10, "A"), T(1, 11, "C") }), "Navigated source tab prevents closing");
        Check(!MergePolicy.SameTabs(source, new List<Tab> { T(1, 10, "A"), T(1, 11, "B"), T(1, 12, "C") }), "Added source tab prevents closing");
        Check(!MergePolicy.SameTabs(new List<Tab>(), new List<Tab>()), "Empty snapshot never authorizes closing");
        Check(MergePolicy.CopiesMatch(source, copies, copies), "Verified copies permit closing");
        Check(!MergePolicy.CopiesMatch(source, copies, new List<Tab> { copies[0] }), "Closed destination tab prevents closing");
        Check(!MergePolicy.CopiesMatch(source, copies, new List<Tab> { copies[0], T(2, 21, "C") }), "Navigated destination tab prevents closing");
        Check(!MergePolicy.CopiesMatch(source, new List<Tab> { copies[0], copies[0] }, copies), "One destination cannot stand in for two tabs");
        Check(!MergePolicy.SameTabs(source, new List<Tab> { T(3, 10, "A"), T(1, 11, "B") }), "Moved source tab prevents closing");
    }
    static IntPtr Open(object shell, string path, Merger merger)
    {
        var before = Native.Windows();
        ((dynamic)shell).Explore(path);
        var timer = Stopwatch.StartNew();
        while (timer.ElapsedMilliseconds < 12000)
        {
            var added = Native.Windows().Except(before).ToList();
            if (added.Count == 1 && merger.Snapshot().Any(t => t.Window == added[0])) return added[0];
            Merger.Pause();
        }
        throw new Exception("Test Explorer window did not open.");
    }
    static void Integration()
    {
        string root = Path.Combine(Path.GetTempPath(), "ExplorerMerge-Test-" + Guid.NewGuid().ToString("N"));
        var paths = new[] { Path.Combine(root, "A"), Path.Combine(root, "B"), Path.Combine(root, "日本語 # 100% & space"), Path.Combine(root, "A") };
        foreach (string path in paths) Directory.CreateDirectory(path);
        var opened = new List<IntPtr>();
        object shell = Activator.CreateInstance(Type.GetTypeFromProgID("Shell.Application"));
        try
        {
            using (var merger = new Merger())
            {
                foreach (string path in paths) opened.Add(Open(shell, path, merger));
                var first = new List<IntPtr> { opened[1], opened[2] };
                merger.Merge(first);
                Check(merger.Errors.Count == 0, "First merge has no errors: " + String.Join("; ", merger.Errors));
                Check(!Native.IsWindow(opened[2]), "Verified source window closes");
                Check(merger.Snapshot().Count(t => t.Window == opened[1]) == 2, "First destination has both tabs");
                merger.Merge(new List<IntPtr> { opened[0], opened[1], opened[3] });
                Check(merger.Errors.Count == 0, "Multi-tab merge has no errors: " + String.Join("; ", merger.Errors));
                var actual = merger.Snapshot().Where(t => t.Window == opened[0]).Select(t => t.Path).OrderBy(p => p).ToArray();
                Check(actual.SequenceEqual(paths.OrderBy(p => p)), "All paths preserved, including Unicode, #, %, spaces and duplicates");
                Check(opened.Skip(1).All(w => !Native.IsWindow(w)), "All verified test source windows closed");
                merger.Merge(new List<IntPtr> { opened[0] });
                Check(merger.Snapshot().Count(t => t.Window == opened[0]) == 4, "Single-window invocation does nothing");
            }
        }
        finally
        {
            // Only windows created by this test are closed. Folders are intentionally
            // left in TEMP, so the test contains no recursive deletion operation.
            foreach (var hwnd in opened) if (Native.IsWindow(hwnd)) Native.PostMessage(hwnd, 0x112, new IntPtr(0xF060), IntPtr.Zero);
            Marshal.ReleaseComObject(shell);
        }
    }
    [STAThread]
    static int Main(string[] args)
    {
        try
        {
            PolicyTests();
            if (args.Contains("--cleanup-test-windows"))
            {
                using (var merger = new Merger())
                    foreach (var group in merger.Snapshot().GroupBy(t => t.Window))
                        if (group.Count() == Native.Tabs(group.Key).Count && group.All(t => t.Path.StartsWith(
                            Path.Combine(Path.GetTempPath(), "ExplorerMerge-Test-"), StringComparison.OrdinalIgnoreCase)))
                            Native.PostMessage(group.Key, 0x112, new IntPtr(0xF060), IntPtr.Zero);
            }
            if (args.Contains("--integration")) Integration();
            Console.WriteLine(checks + " checks passed.");
            return 0;
        }
        catch (Exception e) { Console.Error.WriteLine(e); return 1; }
    }
}

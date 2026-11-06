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
    static void CloseTestWindows(IEnumerable<IntPtr> windows)
    {
        var owned = windows.ToArray();
        foreach (var hwnd in owned) if (Native.IsWindow(hwnd)) Native.PostMessage(hwnd, 0x112, new IntPtr(0xF060), IntPtr.Zero);
        var timer = Stopwatch.StartNew();
        while (owned.Any(Native.IsWindow) && timer.ElapsedMilliseconds < 5000) Merger.Pause();
    }
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
        Process.Start("explorer.exe", "/n,\"" + path + "\"");
        var timer = Stopwatch.StartNew();
        while (timer.ElapsedMilliseconds < 12000)
        {
            var added = Native.Windows().Except(before).ToList();
            string expected = path.StartsWith("shell:", StringComparison.Ordinal) ? path.Substring(6) : path;
            var owned = merger.Snapshot(added).Where(t => t.Path == expected).ToList();
            // Never claim an unrelated window opened by the user during a test.
            if (owned.Count == 1) return owned[0].Window;
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
                var elapsed = Stopwatch.StartNew();
                merger.Merge(first);
                Console.WriteLine("TIMING: 1 tab merge = " + elapsed.ElapsedMilliseconds + " ms");
                Check(merger.Errors.Count == 0, "First merge has no errors: " + String.Join("; ", merger.Errors));
                Check(!Native.IsWindow(opened[2]), "Verified source window closes");
                Check(merger.Snapshot().Count(t => t.Window == opened[1]) == 2, "First destination has both tabs");
                elapsed.Restart();
                merger.Merge(new List<IntPtr> { opened[0], opened[1], opened[3] });
                Console.WriteLine("TIMING: 3 tab merge = " + elapsed.ElapsedMilliseconds + " ms");
                Check(merger.Errors.Count == 0, "Multi-tab merge has no errors: " + String.Join("; ", merger.Errors));
                var actual = merger.Snapshot().Where(t => t.Window == opened[0]).Select(t => t.Path).OrderBy(p => p).ToArray();
                Check(actual.SequenceEqual(paths.OrderBy(p => p)), "All paths preserved, including Unicode, #, %, spaces and duplicates");
                Check(opened.Skip(1).All(w => !Native.IsWindow(w)), "All verified test source windows closed");
                merger.Merge(new List<IntPtr> { opened[0] });
                Check(merger.Snapshot().Count(t => t.Window == opened[0]) == 4, "Single-window invocation does nothing");
                Check(!Native.IsIconic(opened[0]), "Destination restored after successful merge");
            }
        }
        finally
        {
            // Only windows created by this test are closed. Folders are intentionally
            // left in TEMP, so the test contains no recursive deletion operation.
            CloseTestWindows(opened);
            Marshal.ReleaseComObject(shell);
        }
    }
    static void FailureRestoration()
    {
#if TESTING
        string root = Path.Combine(Path.GetTempPath(), "ExplorerMerge-Test-" + Guid.NewGuid().ToString("N"));
        var paths = new[] { Path.Combine(root, "Keep-A"), Path.Combine(root, "Keep-B") };
        foreach (var path in paths) Directory.CreateDirectory(path);
        var opened = new List<IntPtr>();
        object shell = Activator.CreateInstance(Type.GetTypeFromProgID("Shell.Application"));
        try
        {
            using (var merger = new Merger())
            {
                foreach (var path in paths) opened.Add(Open(shell, path, merger));
                Native.ShowWindow(opened[0], 3);
                var original = merger.Snapshot(new[] { opened[1] });
                merger.BeforeCopyVerificationForTest = delegate { throw new InvalidOperationException("Injected verification failure"); };
                merger.Merge(opened);
                Check(merger.Errors.Count == 1 && merger.ClosedWindows == 0, "Verification failure closes no source windows");
                Check(MergePolicy.SameTabs(original, merger.Snapshot(new[] { opened[1] })), "Failure preserves original source tab and location");
                Check(!Native.IsIconic(opened[0]) && Native.IsZoomed(opened[0]), "Failure restores maximized destination");
            }
        }
        finally { CloseTestWindows(opened); Marshal.ReleaseComObject(shell); }
#endif
    }
    static void RecoveryTests()
    {
#if TESTING
        foreach (string scenario in new[] { "transient-start", "unreadable-start", "failed-copy", "transient-final", "source-changed", "copy-changed" })
        {
            string root = Path.Combine(Path.GetTempPath(), "ExplorerMerge-Test-" + Guid.NewGuid().ToString("N"));
            var paths = new[] { "Destination", "Problem", "Healthy", "Changed" }.Select(p => Path.Combine(root, p)).ToArray();
            foreach (var path in paths) Directory.CreateDirectory(path);
            var opened = new List<IntPtr>();
            object shell = Activator.CreateInstance(Type.GetTypeFromProgID("Shell.Application"));
            try
            {
                using (var merger = new Merger())
                {
                    foreach (var path in paths.Take(3)) opened.Add(Open(shell, path, merger));
                    IntPtr problemTab = merger.Snapshot(new[] { opened[1] }).Single().Handle;
                    int unreadableCalls = 0;
                    bool readyWindowClosedBeforeTimeout = false;
                    if (scenario == "transient-start")
                        merger.UnreadableForTest = h => h == problemTab && ++unreadableCalls <= 2;
                    if (scenario == "unreadable-start")
                    {
                        merger.InitialWaitMilliseconds = 250;
                        merger.UnreadableForTest = h => h == problemTab;
                    }
                    if (scenario == "failed-copy")
                    {
                        merger.CopyWaitMilliseconds = 2000;
                        merger.UnverifiedForTest = w => {
                            if (!Native.IsWindow(opened[2])) readyWindowClosedBeforeTimeout = true;
                            return w == opened[1];
                        };
                    }
                    if (scenario == "transient-final")
                        merger.BeforeCopyVerificationForTest = delegate {
                            merger.UnreadableForTest = h => h == problemTab && ++unreadableCalls <= 2;
                        };
                    if (scenario == "source-changed" || scenario == "copy-changed")
                        merger.BeforeCopyVerificationForTest = delegate {
                            IntPtr frame = scenario == "source-changed" ? opened[1] : opened[0];
                            var timer = Stopwatch.StartNew();
                            Tab target = null;
                            while (timer.ElapsedMilliseconds < 5000 && target == null)
                            {
                                target = merger.Snapshot(new[] { frame }).FirstOrDefault(t => t.Path == paths[1]);
                                Merger.Pause();
                            }
                            if (target == null) throw new Exception("Test navigation target not ready.");
                            object windows = ((dynamic)shell).Windows();
                            try
                            {
                                for (int i = 0; i < (int)((dynamic)windows).Count; i++)
                                {
                                    object browser = ((dynamic)windows).Item(i);
                                    try
                                    {
                                        if (Merger.TabHandle(browser) == target.Handle)
                                            ((dynamic)browser).Navigate2(paths[3]);
                                    }
                                    finally { Marshal.ReleaseComObject(browser); }
                                }
                            }
                            finally { Marshal.ReleaseComObject(windows); }
                            timer.Restart();
                            while (!merger.Snapshot(new[] { frame }).Any(t => t.Handle == target.Handle && t.Path == paths[3]))
                            {
                                if (timer.ElapsedMilliseconds > 5000) throw new Exception("Test navigation timed out.");
                                Merger.Pause();
                            }
                            merger.CopyWaitMilliseconds = 1500;
                        };
                    merger.Merge(opened);
                    Check(!Native.IsWindow(opened[2]), scenario + ": healthy source still merges");
                    bool recovered = scenario == "transient-start" || scenario == "transient-final";
                    Check(Native.IsWindow(opened[1]) != recovered, scenario + ": problem source closes only after recovery");
                    Check(merger.Errors.Count == (recovered ? 0 : 1), scenario + ": exact failure count");
                    Check(merger.SkippedWindows == (recovered ? 0 : 1), scenario + ": exact skipped count");
                    var destination = merger.Snapshot(new[] { opened[0] });
                    Check(destination.Any(t => t.Path == paths[2]), scenario + ": healthy copy verified");
                    if (scenario == "failed-copy")
                    {
                        Check(readyWindowClosedBeforeTimeout, "Healthy window finishes while failed copy is still waiting");
                        Check(Program.ResultMessage(merger).Contains("Timed out verifying"), "Dialog includes the specific failure reason");
                    }
                    if (recovered) Check(destination.Any(t => t.Path == paths[1]), scenario + ": recovered copy verified");
                    merger.Merge(new List<IntPtr> { opened[0] });
                    Check(merger.ClosedWindows == 0 && merger.SkippedWindows == 0 && merger.Errors.Count == 0,
                        scenario + ": subsequent no-op clears previous results");
                }
            }
            finally { CloseTestWindows(opened); Marshal.ReleaseComObject(shell); }
        }
#endif
    }
    static void Benchmark()
    {
        string root = Path.Combine(Path.GetTempPath(), "ExplorerMerge-Test-" + Guid.NewGuid().ToString("N"));
        var paths = Enumerable.Range(0, 6).Select(i => Path.Combine(root, "Folder-" + i)).ToArray();
        foreach (string path in paths)
        {
            Directory.CreateDirectory(path);
            for (int i = 0; i < 100; i++) File.WriteAllText(Path.Combine(path, "File-" + i + ".txt"), "test");
        }
        var opened = new List<IntPtr>();
        object shell = Activator.CreateInstance(Type.GetTypeFromProgID("Shell.Application"));
        try
        {
            using (var merger = new Merger())
            {
                foreach (string path in paths) opened.Add(Open(shell, path, merger));
                var elapsed = Stopwatch.StartNew();
                int readsBefore = merger.FolderReads;
                merger.Merge(opened);
                long ms = elapsed.ElapsedMilliseconds;
                int reads = merger.FolderReads - readsBefore;
                Check(merger.Errors.Count == 0, "Benchmark completed: " + String.Join("; ", merger.Errors));
                var actual = merger.Snapshot().Where(t => t.Window == opened[0]).Select(t => t.Path).OrderBy(p => p);
                Check(actual.SequenceEqual(paths.OrderBy(p => p)), "Benchmark preserved all six folders");
                Check(opened.Skip(1).All(w => !Native.IsWindow(w)), "Benchmark closed only verified sources");
                Console.WriteLine("BENCHMARK: 6 windows -> 1, 5 tabs copied, 100 files/folder: " + ms + " ms");
                Console.WriteLine("Folder-location reads during merge: " + reads);
            }
        }
        finally
        {
            CloseTestWindows(opened);
            Marshal.ReleaseComObject(shell);
        }
    }
    static void ShellFolderIntegration()
    {
        string root = Path.Combine(Path.GetTempPath(), "ExplorerMerge-Test-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        object shell = Activator.CreateInstance(Type.GetTypeFromProgID("Shell.Application"));
        var opened = new List<IntPtr>();
        try
        {
            object folder = ((dynamic)shell).NameSpace((int)Environment.SpecialFolder.MyComputer);
            string path;
            try
            {
                object item = ((dynamic)folder).Self;
                try { path = (string)((dynamic)item).Path; }
                finally { Marshal.ReleaseComObject(item); }
            }
            finally { Marshal.ReleaseComObject(folder); }
            using (var merger = new Merger())
            {
                opened.Add(Open(shell, root, merger));
                opened.Add(Open(shell, "shell:" + path, merger));
                merger.Merge(opened);
                Check(merger.Errors.Count == 0, "This PC shell namespace merges successfully");
                Check(merger.Snapshot(new[] { opened[0] }).Any(t => t.Path == path), "This PC location survives PIDL navigation");
                Check(!Native.IsWindow(opened[1]), "Verified shell namespace source closes");
            }
        }
        finally { CloseTestWindows(opened); Marshal.ReleaseComObject(shell); }
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
            if (args.Contains("--integration")) { Integration(); FailureRestoration(); RecoveryTests(); ShellFolderIntegration(); }
            if (args.Contains("--benchmark")) Benchmark();
            Console.WriteLine(checks + " checks passed.");
            return 0;
        }
        catch (Exception e) { Console.Error.WriteLine(e); return 1; }
    }
}

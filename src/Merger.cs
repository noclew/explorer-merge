using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;

namespace ExplorerMerge
{
    sealed class Tab
    {
        internal IntPtr Window, Handle;
        internal string Path;
        internal byte[] Pidl;
    }
    // Conservative gates: a source window is closed only when every original tab
    // still exists unchanged, and every corresponding destination tab is verified.
    static class MergePolicy
    {
        internal static bool SameTabs(IList<Tab> expected, IList<Tab> actual)
        {
            return expected.Count > 0 && expected.Count == actual.Count &&
                expected.All(a => actual.Count(b => a.Handle == b.Handle && a.Window == b.Window &&
                    String.Equals(a.Path, b.Path, StringComparison.Ordinal)) == 1);
        }
        internal static bool CopiesMatch(IList<Tab> source, IList<Tab> copies, IList<Tab> current)
        {
            return source.Count > 0 && source.Count == copies.Count &&
                copies.Select(t => t.Handle).Distinct().Count() == copies.Count &&
                source.Select((s, i) => String.Equals(s.Path, copies[i].Path, StringComparison.Ordinal)).All(b => b) &&
                copies.All(c => current.Any(t => t.Window == c.Window && t.Handle == c.Handle &&
                    String.Equals(t.Path, c.Path, StringComparison.Ordinal)));
        }
    }
    sealed class Merger : IDisposable
    {
        object shell, windows;
        readonly Dictionary<IntPtr, object> browsers = new Dictionary<IntPtr, object>();
        internal int FolderReads;
#if TESTING
        internal Action BeforeCopyVerificationForTest;
        internal Func<IntPtr, bool> UnreadableForTest;
        internal Func<IntPtr, bool> UnverifiedForTest;
        internal int InitialWaitMilliseconds = 2000, CopyWaitMilliseconds = 12000;
#else
        const int InitialWaitMilliseconds = 2000, CopyWaitMilliseconds = 12000;
#endif
        internal int ClosedWindows, CopiedTabs, SkippedWindows;
        internal readonly List<string> Errors = new List<string>();
        internal Merger()
        {
            shell = Activator.CreateInstance(Type.GetTypeFromProgID("Shell.Application", true));
            windows = ((dynamic)shell).Windows();
        }
        static void Release(object value)
        {
            if (value != null && Marshal.IsComObject(value)) Marshal.ReleaseComObject(value);
        }
        internal static IntPtr TabHandle(object browser)
        {
            var sp = (IServiceProvider)browser;
            var guid = new Guid("000214E2-0000-0000-C000-000000000046");
            IShellBrowser sb = null;
            try
            {
                Marshal.ThrowExceptionForHR(sp.QueryService(ref guid, ref guid, out sb));
                IntPtr hwnd;
                Marshal.ThrowExceptionForHR(sb.GetWindow(out hwnd));
                return hwnd;
            }
            finally { Release(sb); }
        }
        Tab Read(object browser, IntPtr handle, bool includePidl)
        {
            if (Native.ClassName(handle) != "ShellTabWindowClass") return null;
#if TESTING
            if (UnreadableForTest != null && UnreadableForTest(handle)) return null;
#endif
            object document = null, folder = null, item = null;
            try
            {
                if ((bool)((dynamic)browser).Busy) return null;
                FolderReads++;
                document = ((dynamic)browser).Document;
                folder = ((dynamic)document).Folder;
                item = ((dynamic)folder).Self;
                string path = (string)((dynamic)item).Path;
                if (String.IsNullOrEmpty(path)) return null;
                return new Tab { Handle = handle, Window = Native.GetAncestor(handle, 2), Path = path,
                    Pidl = includePidl ? Native.ItemIdList(item) : null };
            }
            finally { Release(item); Release(folder); Release(document); }
        }
        internal List<Tab> Snapshot(IEnumerable<IntPtr> onlyWindows = null, bool includePidl = true)
        {
            var result = new List<Tab>();
            var selected = onlyWindows == null ? null : new HashSet<IntPtr>(onlyWindows);
            int count;
            try { count = (int)((dynamic)windows).Count; }
            catch (COMException) { return result; }
            for (int i = 0; i < count; i++)
            {
                object browser = null;
                try
                {
                    browser = ((dynamic)windows).Item(i);
                    if (browser == null) continue;
                    IntPtr handle = TabHandle(browser);
                    if (selected != null && !selected.Contains(Native.GetAncestor(handle, 2))) continue;
                    Tab tab = Read(browser, handle, includePidl);
                    if (tab != null && !result.Any(t => t.Handle == tab.Handle)) result.Add(tab);
                    if (tab != null)
                    {
                        object previous;
                        if (browsers.TryGetValue(handle, out previous)) Release(previous);
                        browsers[handle] = browser;
                        browser = null; // Cache owns this COM reference until Dispose.
                    }
                }
                catch (COMException) { }
                catch (InvalidCastException) { }
                catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException) { }
                finally { Release(browser); }
            }
            return result;
        }
        object FindBrowser(IntPtr handle)
        {
            object cached;
            if (browsers.TryGetValue(handle, out cached)) return cached;
            int count;
            try { count = (int)((dynamic)windows).Count; }
            catch (COMException) { return null; }
            // New ShellWindows entries normally appear at the end. Do not read
            // unrelated folder contents merely to locate a new tab's COM object.
            for (int i = count - 1; i >= 0; i--)
            {
                object browser = null;
                bool matched = false;
                try
                {
                    browser = ((dynamic)windows).Item(i);
                    if (browser == null) continue;
                    try { matched = TabHandle(browser) == handle; }
                    catch (COMException) { }
                    catch (InvalidCastException) { }
                    if (matched)
                    {
                        browsers.Add(handle, browser);
                        cached = browser;
                        browser = null;
                        return cached;
                    }
                }
                catch (COMException) { } // ShellWindows can change during enumeration.
                finally { Release(browser); }
            }
            return null;
        }
        Tab ReadKnown(IntPtr handle, bool includePidl)
        {
            object browser;
            if (!Native.IsWindow(handle) || !browsers.TryGetValue(handle, out browser)) return null;
            try { return Read(browser, handle, includePidl); }
            catch (COMException) { return null; }
            catch (InvalidCastException) { return null; }
            catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException) { return null; }
        }
        List<Tab> ReadKnown(IEnumerable<Tab> tabs)
        {
            return tabs.Select(t => ReadKnown(t.Handle, false)).Where(t => t != null).ToList();
        }
        List<Tab> ReadForVerification(IEnumerable<Tab> tabs)
        {
            var expected = tabs.ToList();
            var timer = Stopwatch.StartNew();
            while (true)
            {
                var current = ReadKnown(expected);
                if (current.Count == expected.Count || timer.ElapsedMilliseconds >= 750 ||
                    expected.Any(t => !Native.IsWindow(t.Handle) || Native.GetAncestor(t.Handle, 2) != t.Window))
                    return current;
                Pause(75);
            }
        }
        internal static void Pause(int milliseconds = 25)
        {
            Application.DoEvents();
            Thread.Sleep(milliseconds);
        }
        static bool Complete(IntPtr window, IList<Tab> tabs)
        {
            var handles = Native.Tabs(window);
            return tabs.Count > 0 && handles.Count == tabs.Count &&
                handles.All(h => tabs.Count(t => t.Handle == h && t.Window == window) == 1);
        }
        sealed class PendingCopy
        {
            internal Tab Source;
            internal IntPtr Handle, Destination;
            internal Tab Verified;
        }
        sealed class SourceGroup
        {
            internal IntPtr Window;
            internal List<Tab> Originals;
            internal readonly List<PendingCopy> Copies = new List<PendingCopy>();
        }

        List<Tab> InitialSnapshot(List<IntPtr> visible)
        {
            var result = Snapshot(visible);
            var timer = Stopwatch.StartNew();
            while (timer.ElapsedMilliseconds < InitialWaitMilliseconds)
            {
                var missing = visible.Where(w => !Complete(w, result.Where(t => t.Window == w).ToList())).ToList();
                if (missing.Count == 0) break;
                Pause(100);
                // Only retry incomplete frames. Do not repeatedly read every folder.
                result.RemoveAll(t => missing.Contains(t.Window));
                result.AddRange(Snapshot(missing));
            }
            return result;
        }

        void RecordFailure(IntPtr window, Exception error)
        {
            string reason = error is COMException
                ? "Explorer communication failed (0x" + error.HResult.ToString("X8") + ")."
                : error.Message;
            Errors.Add(Native.WindowTitle(window) + ": " + reason);
        }
        PendingCopy BeginCopy(Tab source, IntPtr destination)
        {
            var before = Native.Tabs(destination);
            if (before.Count == 0) throw new InvalidOperationException("Destination is unavailable.");
            // Explorer has no documented create-tab API. Never inject global keys.
            if (!Native.PostMessage(before[0], 0x111, new IntPtr(0xA21B), IntPtr.Zero))
                throw new InvalidOperationException("Explorer rejected the new-tab request.");
            var timer = Stopwatch.StartNew();
            IntPtr added = IntPtr.Zero;
            object browser = null;
            while (timer.ElapsedMilliseconds < 8000)
            {
                var newHandles = Native.Tabs(destination).Except(before).ToList();
                if (newHandles.Count > 1) throw new InvalidOperationException("Tabs changed during merging. Please retry without interacting with Explorer.");
                if (newHandles.Count == 1)
                {
                    added = newHandles[0];
                    browser = FindBrowser(added);
                    if (browser != null) break;
                }
                Pause();
            }
            if (added == IntPtr.Zero || browser == null)
                throw new TimeoutException("Explorer did not register the new tab.");
            // A PIDL is passed as a SAFEARRAY of bytes, avoiding URL interpretation
            // of #, %, Unicode, UNC paths and shell namespace folders.
            ((dynamic)browser).Navigate2((object)source.Pidl);
            return new PendingCopy { Source = source, Handle = added, Destination = destination };
        }
        void FinishCopies(List<SourceGroup> groups)
        {
            var timer = Stopwatch.StartNew();
            while (groups.Count > 0)
            {
                foreach (var group in groups.ToArray())
                {
                    foreach (var copy in group.Copies.Where(p => p.Verified == null))
                    {
                        var tab = ReadKnown(copy.Handle, false);
#if TESTING
                        if (UnverifiedForTest != null && UnverifiedForTest(group.Window)) tab = null;
#endif
                        if (tab != null && tab.Window == copy.Destination && tab.Path == copy.Source.Path)
                        {
                            copy.Verified = tab;
                            CopiedTabs++;
                        }
                    }
                    if (group.Copies.Any(p => p.Verified == null)) continue;
                    // A slow or unsupported folder must not block verified windows.
                    try { CloseVerifiedSource(group); }
                    catch (Exception e) { RecordFailure(group.Window, e); }
                    groups.Remove(group);
                }
                if (groups.Count == 0) return;
                if (timer.ElapsedMilliseconds >= CopyWaitMilliseconds) break;
                // Waiting on navigation needs less polling than creating a tab.
                Pause(75);
            }
            foreach (var group in groups)
                RecordFailure(group.Window, new TimeoutException("Timed out verifying " +
                    group.Copies.Count(p => p.Verified == null) + " new tab(s). Original window was kept open."));
        }

        void CloseVerifiedSource(SourceGroup group)
        {
            var copies = group.Copies.Select(p => p.Verified).ToList();
            var current = ReadForVerification(group.Originals.Concat(copies));
            var sourceNow = current.Where(t => t.Window == group.Window).ToList();
            if (!Complete(group.Window, sourceNow) || !MergePolicy.SameTabs(group.Originals, sourceNow) ||
                !MergePolicy.CopiesMatch(group.Originals, copies, current))
                throw new InvalidOperationException("Tabs changed or became unreadable during verification. Original window was kept open.");
            // SC_CLOSE closes the verified frame; automation Quit can close only a tab.
            if (!Native.PostMessage(group.Window, 0x112, new IntPtr(0xF060), IntPtr.Zero))
                throw new InvalidOperationException("Explorer rejected the window-close request.");
            var timer = Stopwatch.StartNew();
            while (Native.IsWindow(group.Window) && timer.ElapsedMilliseconds < 5000) Pause();
            if (Native.IsWindow(group.Window)) throw new TimeoutException("Explorer kept the source window open.");
            foreach (var original in group.Originals)
            {
                object closed;
                if (browsers.TryGetValue(original.Handle, out closed)) Release(closed);
                browsers.Remove(original.Handle);
            }
            ClosedWindows++;
        }
        internal void Merge(List<IntPtr> requested)
        {
            ClosedWindows = CopiedTabs = SkippedWindows = 0;
            Errors.Clear();
            var visible = Native.Windows();
            if (requested != null) visible = requested.Where(h => visible.Contains(h)).Distinct().ToList();
            if (visible.Count < 2) return;
            var initial = InitialSnapshot(visible);
            var eligible = visible.Where(w => Complete(w, initial.Where(t => t.Window == w).ToList())).ToList();
            SkippedWindows += visible.Count - eligible.Count;
            foreach (var window in visible.Except(eligible))
                RecordFailure(window, new InvalidOperationException("Could not read all tabs after waiting for Explorer. Window was kept open."));
            if (eligible.Count < 2) return;
            IntPtr destination = eligible[0];
            bool wasMaximized = Native.IsZoomed(destination);
            // Minimize rather than disable redraw: a crash cannot leave the
            // Explorer frame invisible or permanently unable to repaint.
            if (!Native.IsIconic(destination)) Native.ShowWindow(destination, 7);
            try
            {
                var groups = new List<SourceGroup>();
                foreach (var sourceWindow in eligible.Skip(1))
                {
                    try
                    {
                        var originals = initial.Where(t => t.Window == sourceWindow).ToList();
                        var now = ReadForVerification(originals);
                        if (!Complete(sourceWindow, now) || !MergePolicy.SameTabs(originals, now))
                            throw new InvalidOperationException("Source tabs changed or became unreadable before merging.");
                        var group = new SourceGroup { Window = sourceWindow, Originals = originals };
                        foreach (var original in originals) group.Copies.Add(BeginCopy(original, destination));
                        groups.Add(group);
                    }
                    catch (Exception e) { RecordFailure(sourceWindow, e); }
                }
                // Folder loads overlap; each source is independently verified.
#if TESTING
                if (BeforeCopyVerificationForTest != null) BeforeCopyVerificationForTest();
#endif
                FinishCopies(groups);
            }
            catch (Exception e)
            {
                RecordFailure(destination, e);
            }
            finally
            {
                SkippedWindows += eligible.Skip(1).Count(Native.IsWindow);
                if (Native.IsWindow(destination))
                {
                    Native.ShowWindow(destination, wasMaximized ? 3 : 9);
                    Native.SetForegroundWindow(destination);
                }
            }
        }
        public void Dispose()
        {
            foreach (object browser in browsers.Values) Release(browser);
            browsers.Clear();
            Release(windows);
            Release(shell);
        }
    }
}

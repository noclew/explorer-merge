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
        static Tab Read(object browser)
        {
            IntPtr handle = TabHandle(browser);
            if (Native.ClassName(handle) != "ShellTabWindowClass") return null;
            object document = null, folder = null, item = null;
            try
            {
                if ((bool)((dynamic)browser).Busy) return null;
                document = ((dynamic)browser).Document;
                folder = ((dynamic)document).Folder;
                item = ((dynamic)folder).Self;
                string path = (string)((dynamic)item).Path;
                if (String.IsNullOrEmpty(path)) return null;
                return new Tab { Handle = handle, Window = Native.GetAncestor(handle, 2), Path = path, Pidl = Native.ItemIdList(item) };
            }
            finally { Release(item); Release(folder); Release(document); }
        }
        internal List<Tab> Snapshot()
        {
            var result = new List<Tab>();
            int count = (int)((dynamic)windows).Count;
            for (int i = 0; i < count; i++)
            {
                object browser = null;
                try
                {
                    browser = ((dynamic)windows).Item(i);
                    if (browser == null) continue;
                    Tab tab = Read(browser);
                    if (tab != null && !result.Any(t => t.Handle == tab.Handle)) result.Add(tab);
                }
                catch (COMException) { }
                catch (InvalidCastException) { }
                catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException) { }
                finally { Release(browser); }
            }
            return result;
        }
        void WithBrowser(IntPtr handle, Action<object> action)
        {
            int count = (int)((dynamic)windows).Count;
            for (int i = 0; i < count; i++)
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
                    if (matched) { action(browser); return; }
                }
                finally { Release(browser); }
            }
            throw new InvalidOperationException("Explorer tab disappeared.");
        }
        internal static void Pause()
        {
            Application.DoEvents();
            Thread.Sleep(80);
        }
        static bool Complete(IntPtr window, IList<Tab> tabs)
        {
            var handles = Native.Tabs(window);
            return tabs.Count > 0 && handles.Count == tabs.Count &&
                handles.All(h => tabs.Count(t => t.Handle == h && t.Window == window) == 1);
        }
        Tab Copy(Tab source, IntPtr destination)
        {
            var before = Native.Tabs(destination);
            if (before.Count == 0) throw new InvalidOperationException("Destination is unavailable.");
            // Explorer has no documented create-tab API. Never inject global keys.
            if (!Native.PostMessage(before[0], 0x111, new IntPtr(0xA21B), IntPtr.Zero))
                throw new InvalidOperationException("Explorer rejected the new-tab request.");
            var timer = Stopwatch.StartNew();
            IntPtr added = IntPtr.Zero;
            while (timer.ElapsedMilliseconds < 8000)
            {
                var newHandles = Native.Tabs(destination).Except(before).ToList();
                if (newHandles.Count > 1) throw new InvalidOperationException("Tabs changed during merging. Please retry without interacting with Explorer.");
                if (newHandles.Count == 1)
                {
                    added = newHandles[0];
                    if (Snapshot().Any(t => t.Handle == added)) break;
                }
                Pause();
            }
            if (added == IntPtr.Zero || !Snapshot().Any(t => t.Handle == added))
                throw new TimeoutException("Explorer did not create a readable tab.");
            // A PIDL is passed as a SAFEARRAY of bytes, avoiding URL interpretation
            // of #, %, Unicode, UNC paths and shell namespace folders.
            WithBrowser(added, b => { ((dynamic)b).Navigate2((object)source.Pidl); });
            timer.Restart();
            int stable = 0;
            while (timer.ElapsedMilliseconds < 12000)
            {
                var tab = Snapshot().FirstOrDefault(t => t.Window == destination && t.Handle == added && t.Path == source.Path);
                if (tab != null && ++stable >= 3) { CopiedTabs++; return tab; }
                if (tab == null) stable = 0;
                Pause();
            }
            // Leave both original and destination open on any uncertain outcome.
            throw new TimeoutException("The new tab could not be verified. Original window was kept open.");
        }
        internal void Merge(List<IntPtr> requested)
        {
            var visible = Native.Windows();
            if (requested != null) visible = requested.Where(h => visible.Contains(h)).Distinct().ToList();
            if (visible.Count < 2) return;
            var initial = Snapshot();
            var eligible = visible.Where(w => Complete(w, initial.Where(t => t.Window == w).ToList())).ToList();
            SkippedWindows += visible.Count - eligible.Count;
            if (eligible.Count < 2) return;
            IntPtr destination = eligible[0];
            foreach (var sourceWindow in eligible.Skip(1))
            {
                var originals = initial.Where(t => t.Window == sourceWindow).ToList();
                try
                {
                    var now = Snapshot().Where(t => t.Window == sourceWindow).ToList();
                    if (!Complete(sourceWindow, now) || !MergePolicy.SameTabs(originals, now))
                        throw new InvalidOperationException("Source tabs changed before merging.");
                    var copies = new List<Tab>();
                    foreach (var original in originals) copies.Add(Copy(original, destination));
                    var current = Snapshot();
                    var sourceNow = current.Where(t => t.Window == sourceWindow).ToList();
                    if (!Complete(sourceWindow, sourceNow) || !MergePolicy.SameTabs(originals, sourceNow) ||
                        !MergePolicy.CopiesMatch(originals, copies, current))
                        throw new InvalidOperationException("Tabs changed during verification. Original window was kept open.");
                    // SC_CLOSE is the title-bar Close action. WM_CLOSE and the
                    // shell automation Quit method can close only the active tab.
                    if (!Native.PostMessage(sourceWindow, 0x112, new IntPtr(0xF060), IntPtr.Zero))
                        throw new InvalidOperationException("Explorer rejected the window-close request.");
                    var timer = Stopwatch.StartNew();
                    while (Native.IsWindow(sourceWindow) && timer.ElapsedMilliseconds < 5000) Pause();
                    if (Native.IsWindow(sourceWindow)) throw new TimeoutException("Explorer kept the source window open.");
                    ClosedWindows++;
                }
                catch (Exception e)
                {
                    SkippedWindows++;
                    // Error text never contains folder paths; no telemetry or logs.
                    Errors.Add(e is COMException ? "Explorer communication failed; original window kept." : e.Message);
                    break; // Stop after an uncertain result, avoiding cascading changes.
                }
            }
            if (Native.IsWindow(destination))
            {
                if (Native.IsIconic(destination)) Native.ShowWindow(destination, 9);
                Native.SetForegroundWindow(destination);
            }
        }
        public void Dispose() { Release(windows); Release(shell); }
    }
}

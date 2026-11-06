using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using System.Xml.Linq;
using Microsoft.Win32;

[assembly: System.Reflection.AssemblyTitle("Explorer Merge")]
[assembly: System.Reflection.AssemblyDescription("Merge File Explorer windows into native tabs. No installation or background service.")]
[assembly: System.Reflection.AssemblyVersion("0.3.0.0")]
namespace ExplorerMerge
{
    static class Program
    {
        [STAThread]
        static int Main(string[] args)
        {
            Application.EnableVisualStyles();
            Native.SetCurrentProcessExplicitAppUserModelID("ExplorerMerge.Portable");
            bool acquired;
            using (var mutex = new Mutex(true, @"Local\ExplorerMerge.Portable", out acquired))
            {
                if (!acquired) return 0;
                string report = null;
                try
                {
                    report = Option(args, "--report");
                    string inspect = Option(args, "--inspect");
                    using (var key = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\Windows NT\CurrentVersion"))
                    {
                        int build;
                        if (!Int32.TryParse(Convert.ToString(key.GetValue("CurrentBuildNumber")), out build) || build < 22621)
                            throw new NotSupportedException("Windows 11 22H2 or later is required.");
                    }
                    using (var merger = new Merger())
                    {
                        if (inspect != null)
                        {
                            // Explicit diagnostic export, never enabled during normal use.
                            new XDocument(new XElement("tabs", merger.Snapshot().Select(t => new XElement("tab",
                                new XAttribute("window", t.Window.ToInt64()), new XAttribute("handle", t.Handle.ToInt64()),
                                new XAttribute("path", t.Path))))).Save(inspect);
                            return 0;
                        }
                        string only = Option(args, "--windows");
                        List<IntPtr> requested = only == null ? null : only.Split(',').Select(s => new IntPtr(Int64.Parse(s))).ToList();
                        var elapsed = System.Diagnostics.Stopwatch.StartNew();
                        merger.Merge(requested);
                        elapsed.Stop();
                        if (report != null)
                            new XDocument(new XElement("result", new XAttribute("closedWindows", merger.ClosedWindows),
                                new XAttribute("copiedTabs", merger.CopiedTabs), new XAttribute("skippedWindows", merger.SkippedWindows),
                                new XAttribute("elapsedMilliseconds", elapsed.ElapsedMilliseconds), new XAttribute("folderReads", merger.FolderReads),
                                merger.Errors.Select(e => new XElement("error", e)))).Save(report);
                        if ((merger.Errors.Count > 0 || merger.SkippedWindows > 0) && report == null)
                            MessageBox.Show(ResultMessage(merger),
                                "Explorer Merge", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        return merger.Errors.Count > 0 || merger.SkippedWindows > 0 ? 2 : 0;
                    }
                }
                catch (Exception e)
                {
                    if (report != null)
                    {
                        try { new XDocument(new XElement("result", new XElement("error", e.GetType().Name + ": " + e.Message))).Save(report); }
                        catch (Exception writeError) { ShowError(e.Message + "\n\nCould not write the report: " + writeError.Message); }
                    }
                    else ShowError(e.Message);
                    return 1;
                }
                finally { mutex.ReleaseMutex(); }
            }
        }
        static void ShowError(string message)
        {
            MessageBox.Show("Explorer Merge could not finish.\n\n" + message, "Explorer Merge", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        internal static string ResultMessage(Merger merger)
        {
            return "Merged " + merger.ClosedWindows + " window(s); kept " + merger.SkippedWindows + " window(s) open.\n\n" +
                String.Join("\n", merger.Errors.Take(5)) +
                (merger.Errors.Count > 5 ? "\nAdditional windows could not be merged." : "") +
                "\n\nWait for folders to load and leave Explorer alone during the merge. " +
                "New tabs may already exist; check them before retrying to avoid duplicates.";
        }
        static string Option(string[] args, string name)
        {
            int i = Array.IndexOf(args, name);
            if (i < 0) return null;
            if (i + 1 >= args.Length) throw new ArgumentException("Missing value for " + name);
            return args[i + 1];
        }
    }
}

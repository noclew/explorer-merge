using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;

namespace ExplorerMerge
{
    [ComImport, Guid("6D5140C1-7436-11CE-8034-00AA006009FA"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    interface IServiceProvider
    {
        [PreserveSig] int QueryService(ref Guid service, ref Guid iid, [MarshalAs(UnmanagedType.Interface)] out IShellBrowser browser);
    }
    [ComImport, Guid("000214E2-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    interface IShellBrowser
    {
        [PreserveSig] int GetWindow(out IntPtr window);
    }
    static class Native
    {
        internal delegate bool EnumProc(IntPtr hwnd, IntPtr data);
        [DllImport("user32.dll")] internal static extern bool EnumWindows(EnumProc proc, IntPtr data);
        [DllImport("user32.dll")] internal static extern bool EnumChildWindows(IntPtr parent, EnumProc proc, IntPtr data);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] internal static extern int GetClassName(IntPtr hwnd, StringBuilder text, int count);
        [DllImport("user32.dll")] internal static extern bool IsWindow(IntPtr hwnd);
        [DllImport("user32.dll")] internal static extern bool IsWindowVisible(IntPtr hwnd);
        [DllImport("user32.dll")] internal static extern IntPtr GetAncestor(IntPtr hwnd, uint flags);
        [DllImport("user32.dll", SetLastError = true)] internal static extern bool PostMessage(IntPtr hwnd, uint message, IntPtr wParam, IntPtr lParam);
        [DllImport("user32.dll")] internal static extern bool ShowWindow(IntPtr hwnd, int command);
        [DllImport("user32.dll")] internal static extern bool IsIconic(IntPtr hwnd);
        [DllImport("user32.dll")] internal static extern bool IsZoomed(IntPtr hwnd);
        [DllImport("user32.dll")] internal static extern bool SetForegroundWindow(IntPtr hwnd);
        [DllImport("dwmapi.dll")] static extern int DwmGetWindowAttribute(IntPtr hwnd, int attribute, out int value, int size);
        [DllImport("shell32.dll")] internal static extern int SHGetIDListFromObject([MarshalAs(UnmanagedType.IUnknown)] object obj, out IntPtr pidl);
        [DllImport("shell32.dll")] internal static extern uint ILGetSize(IntPtr pidl);
        [DllImport("shell32.dll", CharSet = CharSet.Unicode)] internal static extern int SetCurrentProcessExplicitAppUserModelID(string id);

        internal static string ClassName(IntPtr hwnd)
        {
            var text = new StringBuilder(256);
            GetClassName(hwnd, text, text.Capacity);
            return text.ToString();
        }
        internal static List<IntPtr> Windows()
        {
            var result = new List<IntPtr>();
            EnumWindows(delegate(IntPtr hwnd, IntPtr data) {
                int cloaked;
                if (ClassName(hwnd) == "CabinetWClass" && IsWindowVisible(hwnd) &&
                    DwmGetWindowAttribute(hwnd, 14, out cloaked, 4) == 0 && cloaked == 0)
                    result.Add(hwnd);
                return true;
            }, IntPtr.Zero);
            return result;
        }
        internal static List<IntPtr> Tabs(IntPtr window)
        {
            var result = new List<IntPtr>();
            EnumChildWindows(window, delegate(IntPtr hwnd, IntPtr data) {
                if (ClassName(hwnd) == "ShellTabWindowClass") result.Add(hwnd);
                return true;
            }, IntPtr.Zero);
            return result;
        }
        internal static byte[] ItemIdList(object folderItem)
        {
            IntPtr pidl;
            Marshal.ThrowExceptionForHR(SHGetIDListFromObject(folderItem, out pidl));
            if (pidl == IntPtr.Zero) throw new InvalidOperationException("Folder has no item ID list.");
            try
            {
                uint size = ILGetSize(pidl);
                if (size < 2 || size > 1048576) throw new InvalidOperationException("Invalid item ID list.");
                var bytes = new byte[size];
                Marshal.Copy(pidl, bytes, 0, bytes.Length);
                return bytes;
            }
            finally { Marshal.FreeCoTaskMem(pidl); }
        }
    }
}

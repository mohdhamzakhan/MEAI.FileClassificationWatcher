using System.Runtime.InteropServices;

namespace MEAI.FileClassificationWatcher
{
    // WinForms' Form.TopMost / Form.Activate() alone don't reliably win against Windows'
    // anti-focus-stealing protection when this popup is created by a background watcher
    // thread while some other app currently has focus — very much the normal case here,
    // since classification prompts fire well after the person's Office app has already
    // closed and they've likely switched to something else in the meantime. Without this,
    // SetForegroundWindow silently degrades to just flashing the taskbar icon instead of
    // actually bringing the window forward, which is exactly the "goes behind" symptom.
    //
    // The fix is the standard documented Win32 workaround: temporarily attach this
    // thread's input to the current foreground window's thread. Windows only blocks
    // SetForegroundWindow calls from threads that don't "own" the current foreground
    // input state — attaching input makes this thread count as if it does, for just long
    // enough to steal focus, then detaches again immediately.
    internal static class ForceForeground
    {
        [DllImport("user32.dll")] private static extern IntPtr GetForegroundWindow();
        [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr hWnd, IntPtr lpdwProcessId);
        [DllImport("kernel32.dll")] private static extern uint GetCurrentThreadId();
        [DllImport("user32.dll")] private static extern bool AttachThreadInput(uint idAttach, uint idAttachTo, bool fAttach);
        [DllImport("user32.dll")] private static extern bool SetForegroundWindow(IntPtr hWnd);
        [DllImport("user32.dll")] private static extern bool BringWindowToTop(IntPtr hWnd);
        [DllImport("user32.dll")] private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
        private const int SW_RESTORE = 9;

        public static void Apply(IntPtr handle)
        {
            try
            {
                var foreground = GetForegroundWindow();
                var foregroundThreadId = GetWindowThreadProcessId(foreground, IntPtr.Zero);
                var currentThreadId = GetCurrentThreadId();
                bool attached = false;

                if (foregroundThreadId != 0 && foregroundThreadId != currentThreadId)
                    attached = AttachThreadInput(currentThreadId, foregroundThreadId, true);

                ShowWindow(handle, SW_RESTORE);
                BringWindowToTop(handle);
                SetForegroundWindow(handle);

                if (attached)
                    AttachThreadInput(currentThreadId, foregroundThreadId, false);
            }
            catch
            {
                // Best-effort — worst case, the window behaves exactly as it did before
                // this fix (visible but possibly not focused), never worse.
            }
        }
    }
}
using System;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;

namespace MEAI.FileClassificationWatcher
{
    // Replaces ACL-based ownership checking for the "is this file mine, or did someone
    // else on the network create it" question. FileInfo.GetAccessControl().GetOwner() is
    // NOT reliable over SMB/network shares — many network filesystems (NAS boxes, non-
    // Windows SMB servers, or Windows shares where IT reassigns ownership for
    // manageability) either report a generic owner for every file, or — the case that was
    // actually happening — the SMB client falls back to reporting YOUR OWN identity as the
    // "owner" of everything on the share, because true per-file owner metadata isn't
    // transmitted. That's what caused prompts for files other users created: the check
    // wasn't failing safe, it was returning "yes, mine" for files that weren't.
    //
    // This uses the Restart Manager API instead (the same mechanism Windows Update/the
    // installer use to find out what needs to close before a file can be replaced) to ask
    // a fundamentally different, more reliable question: "does any process on THIS
    // machine, in THIS interactive session, currently have this file open?" That question
    // doesn't depend on the network filesystem understanding Windows ACLs at all — if
    // another physical user's PC created the file, no process on THIS machine ever had a
    // local handle to it, so the answer is always correctly "no," regardless of what the
    // remote server reports as the file's owner.
    internal static class LocalHandleChecker
    {
        // Retries briefly because the check runs essentially the instant the Created/
        // Changed event fires — if the authoring app briefly hasn't opened its handle yet
        // (rare, but possible under event-dispatch timing), one immediate check could miss it.
        private const int RetryCount = 3;
        private static readonly TimeSpan RetryDelay = TimeSpan.FromMilliseconds(200);

        public static bool IsOpenByCurrentSession(string path)
        {
            for (int attempt = 0; attempt < RetryCount; attempt++)
            {
                if (attempt > 0) Thread.Sleep(RetryDelay);
                if (TryCheck(path, out var openLocally))
                {
                    if (openLocally) return true;
                }
            }
            return false;
        }

        private static bool TryCheck(string path, out bool openByCurrentSession)
        {
            openByCurrentSession = false;
            uint sessionHandle = 0;

            try
            {
                int res = RmStartSession(out sessionHandle, 0, Guid.NewGuid().ToString("N"));
                if (res != 0) return false; // couldn't start a Restart Manager session at all

                string[] resources = { path };
                res = RmRegisterResources(sessionHandle, (uint)resources.Length, resources, 0, null, 0, null);
                if (res != 0) return false;

                uint pnProcInfoNeeded = 0, pnProcInfo = 0, lpdwRebootReasons = 0;
                res = RmGetList(sessionHandle, out pnProcInfoNeeded, ref pnProcInfo, null, ref lpdwRebootReasons);

                if (pnProcInfoNeeded == 0)
                {
                    // Nobody, anywhere, currently has this file open — a definitive answer,
                    // and it means "not open by me" is correct as-is (openByCurrentSession
                    // stays false). The retry loop still gives the authoring app a few more
                    // chances in case it simply hasn't opened its handle yet at this instant.
                    return true;
                }

                var processInfo = new RM_PROCESS_INFO[pnProcInfoNeeded];
                pnProcInfo = pnProcInfoNeeded;
                res = RmGetList(sessionHandle, out pnProcInfoNeeded, ref pnProcInfo, processInfo, ref lpdwRebootReasons);
                if (res != 0) return false;

                var mySessionId = Process.GetCurrentProcess().SessionId;
                openByCurrentSession = processInfo.Take((int)pnProcInfo).Any(p => p.TSSessionId == mySessionId);
                return true;
            }
            catch
            {
                return false; // Restart Manager unavailable for some reason — inconclusive
            }
            finally
            {
                if (sessionHandle != 0) RmEndSession(sessionHandle);
            }
        }

        #region P/Invoke — Restart Manager (rstrtmgr.dll)

        [DllImport("rstrtmgr.dll", CharSet = CharSet.Unicode)]
        private static extern int RmStartSession(out uint pSessionHandle, int dwSessionFlags, string strSessionKey);

        [DllImport("rstrtmgr.dll")]
        private static extern int RmEndSession(uint pSessionHandle);

        [DllImport("rstrtmgr.dll", CharSet = CharSet.Unicode)]
        private static extern int RmRegisterResources(uint pSessionHandle, uint nFiles, string[] rgsFilenames,
            uint nApplications, RM_UNIQUE_PROCESS[]? rgApplications, uint nServices, string[]? rgsServiceNames);

        [DllImport("rstrtmgr.dll")]
        private static extern int RmGetList(uint dwSessionHandle, out uint pnProcInfoNeeded, ref uint pnProcInfo,
            [In, Out] RM_PROCESS_INFO[]? rgAffectedApps, ref uint lpdwRebootReasons);

        [StructLayout(LayoutKind.Sequential)]
        private struct RM_UNIQUE_PROCESS
        {
            public int dwProcessId;
            public System.Runtime.InteropServices.ComTypes.FILETIME ProcessStartTime;
        }

        private const int CCH_RM_MAX_APP_NAME = 255;
        private const int CCH_RM_MAX_SVC_NAME = 63;

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        private struct RM_PROCESS_INFO
        {
            public RM_UNIQUE_PROCESS Process;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = CCH_RM_MAX_APP_NAME + 1)]
            public string strAppName;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = CCH_RM_MAX_SVC_NAME + 1)]
            public string strServiceShortName;
            public RM_APP_TYPE ApplicationType;
            public uint AppStatus;
            public uint TSSessionId;
            [MarshalAs(UnmanagedType.Bool)]
            public bool bRestartable;
        }

        private enum RM_APP_TYPE
        {
            RmUnknownApp = 0,
            RmMainWindow = 1,
            RmOtherWindow = 2,
            RmService = 3,
            RmExplorer = 4,
            RmConsole = 5,
            RmCritical = 1000
        }

        #endregion
    }
}
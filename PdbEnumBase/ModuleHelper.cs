using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace PdbEnum
{
    internal static class ModuleHelper
    {
        // P/Invoke declarations for psapi.dll (Process Status API)
        // Note: These functions are also available in kernel32.dll on Windows 7+,
        // but psapi.dll provides better compatibility
        [DllImport("psapi.dll", SetLastError = true)]
        private static extern bool EnumProcessModules(IntPtr hProcess, [Out] IntPtr[] lphModule, uint cb, out uint lpcbNeeded);

        [DllImport("psapi.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        private static extern uint GetModuleFileNameEx(IntPtr hProcess, IntPtr hModule, StringBuilder lpFilename, uint nSize);

        [DllImport("psapi.dll", SetLastError = true)]
        private static extern bool GetModuleInformation(IntPtr hProcess, IntPtr hModule, out MODULEINFO lpmodinfo, uint cb);

        [StructLayout(LayoutKind.Sequential)]
        private struct MODULEINFO
        {
            public IntPtr lpBaseOfDll;
            public uint SizeOfImage;
            public IntPtr EntryPoint;
        }

        public static ModuleInfo GetModuleInfo(IntPtr processHandle, string moduleName)
        {
            IntPtr hProcess = processHandle;
            IntPtr[] moduleHandles = new IntPtr[1024];

            if (!EnumProcessModules(hProcess, moduleHandles, (uint)(moduleHandles.Length * IntPtr.Size), out uint bytesNeeded))
            {
                int error = Marshal.GetLastWin32Error();
                throw new InvalidOperationException($"EnumProcessModules failed with error: {error}");
            }

            int moduleCount = (int)(bytesNeeded / IntPtr.Size);

            for (int i = 0; i < moduleCount; i++)
            {
                StringBuilder moduleFileName = new StringBuilder(260);
                if (GetModuleFileNameEx(hProcess, moduleHandles[i], moduleFileName, (uint)moduleFileName.Capacity) == 0)
                {
                    continue;
                }

                string fullPath = moduleFileName.ToString();
                string currentModuleName = System.IO.Path.GetFileName(fullPath);

                if (string.Equals(currentModuleName, moduleName, StringComparison.OrdinalIgnoreCase))
                {
                    if (GetModuleInformation(hProcess, moduleHandles[i], out MODULEINFO modInfo, (uint)Marshal.SizeOf<MODULEINFO>()))
                    {
                        return new ModuleInfo
                        {
                            Name = currentModuleName,
                            FullPath = fullPath,
                            BaseAddress = (ulong)modInfo.lpBaseOfDll.ToInt64(),
                            Size = modInfo.SizeOfImage,
                            EntryPoint = (ulong)modInfo.EntryPoint.ToInt64()
                        };
                    }
                }
            }

            return null;
        }
    }
}

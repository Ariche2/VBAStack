using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace PdbEnum
{
    internal class SymbolEnumerator
    {
        private readonly IntPtr _pHandle;
        private ulong _moduleBase;
        private bool _symbolsLoaded;
        private static bool _quietMode = false;

        // P/Invoke declarations for DbgHelp.dll

        [DllImport(("dbghelp_amd64.dll"), EntryPoint = "SymInitialize", SetLastError = true, CharSet = CharSet.Ansi)]
        private static extern bool amd64_SymInitialize(IntPtr hProcess, string UserSearchPath, [MarshalAs(UnmanagedType.Bool)] bool fInvadeProcess);

        [DllImport("dbghelp_amd64.dll", EntryPoint = "SymGetOptions", SetLastError = true)]
        private static extern uint amd64_SymGetOptions();

        [DllImport("dbghelp_amd64.dll", EntryPoint = "SymSetOptions", SetLastError = true)]
        private static extern uint amd64_SymSetOptions(uint SymOptions);

        [DllImport("dbghelp_amd64.dll", EntryPoint = "SymLoadModuleEx", SetLastError = true, CharSet = CharSet.Ansi)]
        private static extern ulong amd64_SymLoadModuleEx(IntPtr hProcess, IntPtr hFile, string ImageName, string ModuleName, ulong BaseOfDll, uint DllSize, IntPtr Data, uint Flags);

        [DllImport("dbghelp_amd64.dll", EntryPoint = "SymGetModuleBase64", SetLastError = true)]
        private static extern ulong amd64_SymGetModuleBase64(IntPtr hProcess, ulong dwAddr);

        [DllImport("dbghelp_amd64.dll", EntryPoint = "SymGetModuleInfo64", SetLastError = true, CharSet = CharSet.Ansi)]
        private static extern bool amd64_SymGetModuleInfo64(IntPtr hProcess, ulong dwAddr, ref IMAGEHLP_MODULE64 ModuleInfo);

        [DllImport("dbghelp_amd64.dll", EntryPoint = "SymEnumSymbols", SetLastError = true, CharSet = CharSet.Ansi)]
        private static extern bool amd64_SymEnumSymbols(IntPtr hProcess, ulong BaseOfDll, string Mask, SymEnumSymbolsProc EnumSymbolsCallback, IntPtr UserContext);

        [DllImport("dbghelp_amd64.dll", EntryPoint = "SymUnloadModule64", SetLastError = true)]
        private static extern bool amd64_SymUnloadModule64(IntPtr hProcess, ulong BaseOfDll);

        [DllImport("dbghelp_amd64.dll", EntryPoint = "SymCleanup", SetLastError = true)]
        private static extern bool amd64_SymCleanup(IntPtr hProcess);

        [DllImport(("dbghelp_x86.dll"), EntryPoint = "SymInitialize", SetLastError = true, CharSet = CharSet.Ansi)]
        private static extern bool x86_SymInitialize(IntPtr hProcess, string UserSearchPath, [MarshalAs(UnmanagedType.Bool)] bool fInvadeProcess);

        [DllImport("dbghelp_x86.dll", EntryPoint = "SymGetOptions", SetLastError = true)]
        private static extern uint x86_SymGetOptions();

        [DllImport("dbghelp_x86.dll", EntryPoint = "SymSetOptions", SetLastError = true)]
        private static extern uint x86_SymSetOptions(uint SymOptions);

        [DllImport("dbghelp_x86.dll", EntryPoint = "SymLoadModuleEx", SetLastError = true, CharSet = CharSet.Ansi)]
        private static extern ulong x86_SymLoadModuleEx(IntPtr hProcess, IntPtr hFile, string ImageName, string ModuleName, ulong BaseOfDll, uint DllSize, IntPtr Data, uint Flags);

        [DllImport("dbghelp_x86.dll", EntryPoint = "SymGetModuleBase64", SetLastError = true)]
        private static extern ulong x86_SymGetModuleBase64(IntPtr hProcess, ulong dwAddr);

        [DllImport("dbghelp_x86.dll", EntryPoint = "SymGetModuleInfo64", SetLastError = true, CharSet = CharSet.Ansi)]
        private static extern bool x86_SymGetModuleInfo64(IntPtr hProcess, ulong dwAddr, ref IMAGEHLP_MODULE64 ModuleInfo);

        [DllImport("dbghelp_x86.dll", EntryPoint = "SymEnumSymbols", SetLastError = true, CharSet = CharSet.Ansi)]
        private static extern bool x86_SymEnumSymbols(IntPtr hProcess, ulong BaseOfDll, string Mask, SymEnumSymbolsProc EnumSymbolsCallback, IntPtr UserContext);

        [DllImport("dbghelp_x86.dll", EntryPoint = "SymUnloadModule64", SetLastError = true)]
        private static extern bool x86_SymUnloadModule64(IntPtr hProcess, ulong BaseOfDll);

        [DllImport("dbghelp_x86.dll", EntryPoint = "SymCleanup", SetLastError = true)]
        private static extern bool x86_SymCleanup(IntPtr hProcess);

        private delegate bool SymEnumSymbolsProc(IntPtr pSymInfo, uint SymbolSize, IntPtr UserContext);

        // Symbol options
        private const uint SYMOPT_UNDNAME = 0x00000002;
        private const uint SYMOPT_LOAD_LINES = 0x00000010;
        private const uint SYMOPT_FAIL_CRITICAL_ERRORS = 0x00000200;
        private const uint SYMOPT_ALLOW_ABSOLUTE_SYMBOLS = 0x00000400;
        private const uint SYMOPT_INCLUDE_32BIT_MODULES = 0x00002000;
        private const uint SYMOPT_DEBUG = 0x80000000;

        public SymbolEnumerator(IntPtr processHandle)
        {
            _pHandle = processHandle;
        }

        public static void SetQuietMode(bool quiet)
        {
            _quietMode = quiet;
        }

        private static void Log(string message)
        {
            if (!_quietMode)
            {
                Console.Error.WriteLine(message);
            }
        }

        public bool InitializeSymbols(string symbolPath = null)
        {
            if (string.IsNullOrEmpty(symbolPath))
            {
                // Use AppData\Local\Temp for symbol cache
                string tempPath = System.IO.Path.GetTempPath();
                string cacheDir = System.IO.Path.Combine(tempPath, "PdbEnum_Symbols");

                symbolPath = $"SRV*{cacheDir}*https://msdl.microsoft.com/download/symbols";

                // Ensure the symbol cache directory exists
                if (!System.IO.Directory.Exists(cacheDir))
                {
                    try
                    {
                        System.IO.Directory.CreateDirectory(cacheDir);
                        Log($"[Debug] Created symbol cache directory: {cacheDir}");
                    }
                    catch (Exception ex)
                    {
                        Log($"[Warning] Failed to create symbol cache directory: {ex.Message}");
                        // Fallback to using symbol server without local cache
                        symbolPath = "https://msdl.microsoft.com/download/symbols";
                    }
                }
            }


            Log($"[Debug] Symbol path: {symbolPath}");

            bool result = SymInitialize(_pHandle, symbolPath, false);
            if (!result)
            {
                int error = Marshal.GetLastWin32Error();
                throw new InvalidOperationException($"SymInitialize failed with error: {error}");
            }

            uint options = SymGetOptions();
            options |= SYMOPT_UNDNAME;
            options |= SYMOPT_LOAD_LINES;
            options |= SYMOPT_FAIL_CRITICAL_ERRORS;
            options |= SYMOPT_ALLOW_ABSOLUTE_SYMBOLS;
            options |= SYMOPT_INCLUDE_32BIT_MODULES;
            options |= SYMOPT_DEBUG;

            SymSetOptions(options);
            Log($"[Debug] Symbol options set: 0x{options:X}");

            return result;
        }

        public bool LoadModule(string modulePath, ulong baseAddress, uint size)
        {
            Log($"[Debug] Loading module: {modulePath}");
            Log($"[Debug] Base address: 0x{baseAddress:X}");
            Log($"[Debug] Size: {size} bytes");

            _moduleBase = SymLoadModuleEx(
                _pHandle,
                IntPtr.Zero,
                modulePath,
                null,
                baseAddress,
                size,
                IntPtr.Zero,
                0);

            if (_moduleBase == 0)
            {
                int error = Marshal.GetLastWin32Error();
                Log($"[Debug] SymLoadModuleEx returned 0, Win32Error: {error}");

                // Error code 0 means the module was already loaded (e.g., by SymInitialize with fInvadeProcess=true)
                if (error == 0)
                {
                    _moduleBase = SymGetModuleBase64(_pHandle, baseAddress);
                    if (_moduleBase == 0)
                    {
                        throw new InvalidOperationException($"Module at base address 0x{baseAddress:X} is not loaded in the symbol handler.");
                    }
                    Log($"[Debug] Module was already loaded at base: 0x{_moduleBase:X}");
                }
                else
                {
                    throw new InvalidOperationException($"SymLoadModuleEx failed with error: {error}");
                }
            }

            Log($"[Debug] Module loaded at base: 0x{_moduleBase:X}");

            IMAGEHLP_MODULE64 moduleInfo = new()
            {
                SizeOfStruct = (uint)Marshal.SizeOf<IMAGEHLP_MODULE64>()
            };

            if (SymGetModuleInfo64(_pHandle, _moduleBase, ref moduleInfo))
            {
                Log($"[Debug] Symbol type: {moduleInfo.SymType}");
                Log($"[Debug] Loaded image: {moduleInfo.LoadedImageName}");
                Log($"[Debug] Loaded PDB: {moduleInfo.LoadedPdbName}");

                if (moduleInfo.SymType == 4) // SymExport
                {
                    Log("[WARNING] Only exports loaded - PDB not found or failed to load!");
                    Log("[WARNING] Check the debug output above for details.");
                }
            }

            _symbolsLoaded = true;
            return true;
        }

        public PdbInfo GetPdbInfo()
        {
            if (!_symbolsLoaded)
            {
                throw new InvalidOperationException("Module not loaded. Call LoadModule first.");
            }

            IMAGEHLP_MODULE64 moduleInfo = new()
            {
                SizeOfStruct = (uint)Marshal.SizeOf<IMAGEHLP_MODULE64>()
            };

            if (!SymGetModuleInfo64(_pHandle, _moduleBase, ref moduleInfo))
            {
                int error = Marshal.GetLastWin32Error();
                throw new InvalidOperationException($"SymGetModuleInfo64 failed with error: {error}");
            }

            return new PdbInfo
            {
                PdbGuid = moduleInfo.PdbSig70,
                PdbAge = moduleInfo.PdbAge,
                PdbFileName = moduleInfo.LoadedPdbName,
                SymType = moduleInfo.SymType
            };
        }

        public SymbolInfo FindSymbol(string symbolName)
        {
            if (!_symbolsLoaded)
            {
                throw new InvalidOperationException("Module not loaded. Call LoadModule first.");
            }

            SymbolInfo foundSymbol = null;
            string targetSymbol = symbolName.ToLowerInvariant();

            bool callback(IntPtr symInfo, uint symbolSize, IntPtr userContext)
            {
                if (symInfo != IntPtr.Zero)
                {
                    SYMBOL_INFO _symInfo = Marshal.PtrToStructure<SYMBOL_INFO>(symInfo);
                    if (_symInfo.Name != null && _symInfo.Name.ToLowerInvariant().Contains(targetSymbol))
                    {
                        foundSymbol = new SymbolInfo
                        {
                            Name = _symInfo.Name,
                            Address = _symInfo.Address,
                            Size = _symInfo.Size,
                            Flags = (uint)_symInfo.Flags,
                            Tag = (uint)_symInfo.Tag
                        };
                        return false;
                    }
                }
                return true;
            }

            if (!SymEnumSymbols(
                _pHandle,
                _moduleBase,
                "*",
                callback,
                IntPtr.Zero))
            {
                int error = Marshal.GetLastWin32Error();
                Log($"[Debug] SymEnumSymbols failed with error: {error}");
            }

            return foundSymbol;
        }

        public List<SymbolInfo> FindSymbolsAtAddresses(List<ulong> addresses)
        {
            if (!_symbolsLoaded)
            {
                throw new InvalidOperationException("Module not loaded. Call LoadModule first.");
            }
            int symbolsToFind = addresses.Count;
            int foundSymbols = 0;
            List<SymbolInfo> retList = new();
            foreach (ulong addr in addresses)
            {
                retList.Add(new SymbolInfo
                {
                    Name = string.Empty,
                    Address = addr,
                    Size = 0,
                    Flags = 0,
                    Tag = 0
                });
            }

            bool callback(IntPtr symInfo, uint symbolSize, IntPtr userContext)
            {
                if (symInfo != IntPtr.Zero)
                {
                    SYMBOL_INFO _symInfo = Marshal.PtrToStructure<SYMBOL_INFO>(symInfo);
                    //Log("[Debug] Checking symbol: " + _symInfo.Name + $" at address 0x{_symInfo.Address:X}");
                    foreach (SymbolInfo retSym in retList)
                    {
                        if (_symInfo.Name != null && _symInfo.Address == retSym.Address)
                        {
                            retSym.Name = _symInfo.Name;
                            retSym.Address = _symInfo.Address;
                            retSym.Size = _symInfo.Size;
                            retSym.Flags = (uint)_symInfo.Flags;
                            retSym.Tag = (uint)_symInfo.Tag;
                            foundSymbols++;
                            if (foundSymbols >= symbolsToFind)
                            {
                                return false;
                            }
                        }

                    }
                }
                return true;
            }

            if (!SymEnumSymbols(
                _pHandle,
                _moduleBase,
                "*",
                callback,
                IntPtr.Zero))
            {
                int error = Marshal.GetLastWin32Error();
                Log($"[Debug] SymEnumSymbols failed with error: {error}");
            }

            return retList;
        }
        public List<SymbolInfo> EnumerateAllSymbols(string filter = "*")
        {
            if (!_symbolsLoaded)
            {
                throw new InvalidOperationException("Module not loaded. Call LoadModule first.");
            }

            List<SymbolInfo> symbols = new();

            bool callback(IntPtr symInfo, uint symbolSize, IntPtr userContext)
            {
                if (symInfo != IntPtr.Zero)
                {
                    SYMBOL_INFO _symInfo = Marshal.PtrToStructure<SYMBOL_INFO>(symInfo);
                    symbols.Add(new SymbolInfo
                    {
                        Name = _symInfo.Name,
                        Address = _symInfo.Address,
                        Size = _symInfo.Size,
                        Flags = (uint)_symInfo.Flags,
                        Tag = (uint)_symInfo.Tag
                    });
                }
                return true;
            }

            if (!SymEnumSymbols(
                _pHandle,
                _moduleBase,
                filter,
                callback,
                IntPtr.Zero))
            {
                int error = Marshal.GetLastWin32Error();
                Log($"[Debug] SymEnumSymbols failed with error: {error}");
            }

            return symbols;
        }

        public void Cleanup()
        {
            if (_symbolsLoaded && _moduleBase != 0)
            {
                SymUnloadModule64(_pHandle, _moduleBase);
                _symbolsLoaded = false;
            }

            SymCleanup(_pHandle);
        }

        private void SymCleanup(IntPtr pHandle)
        {
            if (IntPtr.Size == 8)
            {
                amd64_SymCleanup(pHandle);
            }
            else
            {
                x86_SymCleanup(pHandle);
            }
        }

        private void SymUnloadModule64(IntPtr pHandle, ulong moduleBase)
        {
            if (IntPtr.Size == 8)
            {
                amd64_SymUnloadModule64(pHandle, moduleBase);
            }
            else
            {
                x86_SymUnloadModule64(pHandle, moduleBase);
            }
        }
        private void SymSetOptions(uint options)
        {
            if (IntPtr.Size == 8)
            {
                amd64_SymSetOptions(options);
            }
            else
            {
                x86_SymSetOptions(options);
            }
        }

        private uint SymGetOptions()
        {
            if (IntPtr.Size == 8)
            {
                return amd64_SymGetOptions();
            }
            else
            {
                return x86_SymGetOptions();
            }
        }

        private bool SymInitialize(IntPtr pHandle, string symbolPath, bool v)
        {
            if (IntPtr.Size == 8)
            {
                return amd64_SymInitialize(pHandle, symbolPath, v);
            }
            else
            {
                return x86_SymInitialize(pHandle, symbolPath, v);
            }
        }

        private bool SymGetModuleInfo64(IntPtr pHandle, ulong moduleBase, ref IMAGEHLP_MODULE64 moduleInfo)
        {
            if (IntPtr.Size == 8)
            {
                return amd64_SymGetModuleInfo64(pHandle, moduleBase, ref moduleInfo);
            }
            else
            {
                return x86_SymGetModuleInfo64(pHandle, moduleBase, ref moduleInfo);
            }
        }

        private ulong SymGetModuleBase64(IntPtr pHandle, ulong baseAddress)
        {
            if (IntPtr.Size == 8)
            {
                return amd64_SymGetModuleBase64(pHandle, baseAddress);
            }
            else
            {
                return x86_SymGetModuleBase64(pHandle, baseAddress);

            }
        }

        private ulong SymLoadModuleEx(IntPtr pHandle, IntPtr zero1, string modulePath, object value, ulong baseAddress, uint size, IntPtr zero2, int v)
        {
            if (IntPtr.Size == 8)
            {
                return amd64_SymLoadModuleEx(pHandle, zero1, modulePath, null, baseAddress, size, zero2, (uint)v);
            }
            else
            {
                return x86_SymLoadModuleEx(pHandle, zero1, modulePath, null, baseAddress, size, zero2, (uint)v);
            }
        }

        private bool SymEnumSymbols(IntPtr pHandle, ulong moduleBase, string v, SymEnumSymbolsProc callback, IntPtr zero)
        {
            if (IntPtr.Size == 8)
            {
                return amd64_SymEnumSymbols(pHandle, moduleBase, v, callback, zero);
            }
            else
            {
                return x86_SymEnumSymbols(pHandle, moduleBase, v, callback, zero);
            }
        }


    }
}

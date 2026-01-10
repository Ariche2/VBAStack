using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using PdbEnum;

#if (X64) 
namespace PdbEnum_x64
#elif (X86)
namespace PdbEnum_x86
#endif
{
    internal class Program
    {
        // P/Invoke declarations for Kernel32.dll
        [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        private static extern IntPtr LoadLibraryEx(string lpFileName, IntPtr hFile, uint dwFlags);

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern IntPtr OpenProcess(uint dwDesiredAccess, bool bInheritHandle, uint dwProcessId);

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern bool CloseHandle(IntPtr hObject);

        private const uint PROCESS_QUERY_INFORMATION = 0x0400;
        private const uint PROCESS_VM_READ = 0x0010;

        static void Main(string[] args)
        {
            OutputFormat outputFormat = OutputFormat.Human;
            bool quietMode = false;
            int argOffset = 0;

            // Parse optional flags
            while (argOffset < args.Length && args[argOffset].StartsWith("-"))
            {
                string flag = args[argOffset].ToLowerInvariant();
                switch (flag)
                {
                    case "-json":
                        outputFormat = OutputFormat.Json;
                        break;
                    case "-xml":
                        outputFormat = OutputFormat.Xml;
                        break;
                    case "-q":
                    case "-quiet":
                        quietMode = true;
                        break;
                    default:
                        Console.Error.WriteLine($"Unknown flag: {args[argOffset]}");
                        ShowUsage();
                        return;
                }
                argOffset++;
            }

            if (args.Length - argOffset < 3)
            {
                ShowUsage();
                return;
            }

            if (!int.TryParse(args[argOffset], out int processId))
            {
                Console.Error.WriteLine($"Error: Invalid process ID '{args[argOffset]}'");
                ShowUsage();
                return;
            }

            string moduleName = args[argOffset + 1];

            // Collect all remaining arguments as symbol names
            List<string> symbolNames = new();
            for (int i = argOffset + 2; i < args.Length; i++)
            {
                symbolNames.Add(args[i]);
            }

            try
            {
                FindSymbols(processId, moduleName, symbolNames, outputFormat, quietMode);
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
                if (ex.InnerException != null)
                {
                    Console.Error.WriteLine($"Inner Exception: {ex.InnerException.Message}");
                }
                Environment.Exit(1);
            }
        }

        static void ShowUsage()
        {
            Console.WriteLine("PdbEnum - Symbol Enumerator using DbgHelp");
            Console.WriteLine();
            Console.WriteLine("Usage: PdbEnum.exe [options] <ProcessID> <ModuleName> <SymbolName> [<SymbolName2> ...]");
            Console.WriteLine();
            Console.WriteLine("Options:");
            Console.WriteLine("  -json       Output in JSON format");
            Console.WriteLine("  -xml        Output in XML format");
            Console.WriteLine("  -q, -quiet  Suppress informational messages (useful with structured output)");
            Console.WriteLine();
            Console.WriteLine("Arguments:");
            Console.WriteLine("  ProcessID   - The ID of the process containing the module");
            Console.WriteLine("  ModuleName  - The name of the module (e.g., ntdll.dll)");
            Console.WriteLine("  SymbolName  - One or more symbol names to search for (case-insensitive, partial match)");
            Console.WriteLine();
            Console.WriteLine("Examples:");
            Console.WriteLine("  PdbEnum.exe 1234 ntdll.dll NtCreateFile");
            Console.WriteLine("  PdbEnum.exe -json 1234 ntdll.dll NtCreateFile");
            Console.WriteLine("  PdbEnum.exe -xml -quiet 1234 kernel32.dll CreateFile");
            Console.WriteLine("  PdbEnum.exe -json 1234 VBE7.DLL EbMode EbSetMode EbGetCallstackCount");
            Console.WriteLine();
            Console.WriteLine("Symbol Path:");
            Console.WriteLine("  Default: SRV*C:\\Symbols*https://msdl.microsoft.com/download/symbols");
            Console.WriteLine("  Set _NT_SYMBOL_PATH environment variable to override");
        }

        static void FindSymbol(int processId, string moduleName, string symbolName, OutputFormat outputFormat, bool quietMode)
        {
            TextWriter log = quietMode ? TextWriter.Null : Console.Error;
            SymbolEnumerator.SetQuietMode(quietMode);

            IntPtr hinstDbgHelp;
            string dbghelpPath = Process.GetCurrentProcess().MainModule.FileName.Replace(Path.GetFileName(Process.GetCurrentProcess().MainModule.FileName), "runtimes");
            if (IntPtr.Size == 8)
            {
                dbghelpPath += "\\win-x64\\dbghelp.dll";
            }
            else
            {
                dbghelpPath += "\\win-x86\\dbghelp.dll";
            }

            if (File.Exists(dbghelpPath))
            {
                log.WriteLine($"Loading dbghelp.dll from {dbghelpPath}...");
                hinstDbgHelp = LoadLibraryEx(dbghelpPath, IntPtr.Zero, 0);
                if (hinstDbgHelp == IntPtr.Zero)
                {
                    int err = Marshal.GetLastWin32Error();
                    throw new InvalidOperationException($"Failed to load dbghelp.dll from {dbghelpPath}. Error: {err}");
                }
            }
            else if (File.Exists(dbghelpPath.Replace(".dll", ".notdll")))
            {
                File.Copy(dbghelpPath.Replace(".dll", ".notdll"), dbghelpPath, true);
                File.Copy(dbghelpPath.Replace(".dll", ".notdll").Replace("DbgHelp", "symsrv"), dbghelpPath.Replace("dbghelp.dll", "symsrv.dll"), true);
                log.WriteLine($"Loading dbghelp.dll from {dbghelpPath}...");
                hinstDbgHelp = LoadLibraryEx(dbghelpPath, IntPtr.Zero, 0);
                if (hinstDbgHelp == IntPtr.Zero)
                {
                    int err = Marshal.GetLastWin32Error();
                    throw new InvalidOperationException($"Failed to load dbghelp.dll from {dbghelpPath}. Error: {err}");
                }
            }
            else
            {
                //Fail. ONLY use the DBGHelp.dll we provide
                throw new InvalidOperationException("Failed to find dbghelp.dll in the expected location. Make sure the DbgHelp folder with the appropriate architecture subfolder exists alongside PdbEnum.exe.");
            }


            log.WriteLine($"Opening process {processId}...");
            IntPtr processHandle = OpenProcess(
                PROCESS_QUERY_INFORMATION | PROCESS_VM_READ,
                false,
                (uint)processId);

            if (processHandle == IntPtr.Zero)
            {
                int error = Marshal.GetLastWin32Error();
                throw new InvalidOperationException($"Failed to open process {processId}. Error: {error}. Make sure you have appropriate permissions.");
            }

            SymbolSearchResult result = new() { Success = false };

            try
            {
                log.WriteLine($"Finding module '{moduleName}'...");
                ModuleInfo moduleInfo = ModuleHelper.GetModuleInfo(processHandle, moduleName) ?? throw new InvalidOperationException($"Module '{moduleName}' not found in process {processId}");
                result.Module = moduleInfo;

                if (outputFormat == OutputFormat.Human)
                {
                    log.WriteLine(moduleInfo.ToString());
                    log.WriteLine();
                }

                string symbolPath = Environment.GetEnvironmentVariable("_NT_SYMBOL_PATH");
                if (string.IsNullOrEmpty(symbolPath))
                {
                    symbolPath = "SRV*C:\\Symbols*https://msdl.microsoft.com/download/symbols";
                }

                log.WriteLine($"Symbol Path: {symbolPath}");
                log.WriteLine();

                SymbolEnumerator enumerator = new(processHandle);

                try
                {
                    log.WriteLine("Initializing symbol handler...");
                    enumerator.InitializeSymbols(symbolPath);

                    log.WriteLine($"Loading symbols for '{moduleInfo.Name}'...");
                    log.WriteLine("(This may take a while if PDB needs to be downloaded)");
                    enumerator.LoadModule(moduleInfo.FullPath, moduleInfo.BaseAddress, moduleInfo.Size);

                    PdbInfo pdbInfo = enumerator.GetPdbInfo();
                    result.PdbInfo = pdbInfo;

                    if (outputFormat == OutputFormat.Human)
                    {
                        log.WriteLine();
                        log.WriteLine(pdbInfo.ToString());
                        log.WriteLine();
                    }

                    log.WriteLine($"Searching for symbol containing '{symbolName}'...");
                    log.WriteLine();

                    SymbolInfo symbol = enumerator.FindSymbol(symbolName);
                    result.Symbol = symbol;
                    result.SearchedSymbolName = symbolName;
                    result.Success = true;

                    OutputFormatter.WriteResult(result, outputFormat, Console.Out);

                    if (symbol == null && outputFormat == OutputFormat.Human)
                    {
                        log.WriteLine();
                        log.WriteLine("Tip: The search is case-insensitive and matches partial names.");
                        log.WriteLine("      Try a shorter or different part of the symbol name.");
                    }
                }
                catch (Exception ex)
                {
                    log.WriteLine(ex.ToString());
                    throw;
                }
                finally
                {
                    enumerator.Cleanup();
                }
            }
            finally
            {
                if (processHandle != IntPtr.Zero)
                {
                    CloseHandle(processHandle);
                }
            }
        }

        static void FindSymbols(int processId, string moduleName, List<string> symbolNames, OutputFormat outputFormat, bool quietMode)
        {
            TextWriter log = quietMode ? TextWriter.Null : Console.Error;
            SymbolEnumerator.SetQuietMode(quietMode);

            IntPtr hinstDbgHelp;
            string dbghelpPath = Process.GetCurrentProcess().MainModule.FileName.Replace(Path.GetFileName(Process.GetCurrentProcess().MainModule.FileName), "runtimes");
            if (IntPtr.Size == 8)
            {
                dbghelpPath += "\\win-x64\\dbghelp.dll";
            }
            else
            {
                dbghelpPath += "\\win-x86\\dbghelp.dll";
            }

            if (File.Exists(dbghelpPath))
            {
                log.WriteLine($"Loading dbghelp.dll from {dbghelpPath}...");
                hinstDbgHelp = LoadLibraryEx(dbghelpPath, IntPtr.Zero, 0);
                if (hinstDbgHelp == IntPtr.Zero)
                {
                    int err = Marshal.GetLastWin32Error();
                    throw new InvalidOperationException($"Failed to load dbghelp.dll from {dbghelpPath}. Error: {err}");
                }
            }
            else if (File.Exists(dbghelpPath.Replace(".dll", ".notdll"))) //this is part of an ugly hack to make ClickOnce/VSTO stop complaining about these DLLs being assemblies
            {
                File.Copy(dbghelpPath.Replace(".dll", ".notdll"), dbghelpPath, true);
                File.Copy(dbghelpPath.Replace(".dll", ".notdll").Replace("DbgHelp", "symsrv"), dbghelpPath.Replace("dbghelp.dll", "symsrv.dll"), true);
                log.WriteLine($"Loading dbghelp.dll from {dbghelpPath}...");
                hinstDbgHelp = LoadLibraryEx(dbghelpPath, IntPtr.Zero, 0);
                if (hinstDbgHelp == IntPtr.Zero)
                {
                    int err = Marshal.GetLastWin32Error();
                    throw new InvalidOperationException($"Failed to load dbghelp.dll from {dbghelpPath}. Error: {err}");
                }
            }
            else
            {
                //Fail. ONLY use the DBGHelp.dll we provide
                throw new InvalidOperationException("Failed to find dbghelp.dll in the expected location. Make sure the DbgHelp folder with the appropriate architecture subfolder exists alongside PdbEnum.exe.");
            }

            log.WriteLine($"Opening process {processId}...");
            IntPtr processHandle = OpenProcess(
                PROCESS_QUERY_INFORMATION | PROCESS_VM_READ,
                false,
                (uint)processId);

            if (processHandle == IntPtr.Zero)
            {
                int error = Marshal.GetLastWin32Error();
                throw new InvalidOperationException($"Failed to open process {processId}. Error: {error}. Make sure you have appropriate permissions.");
            }

            BatchSymbolSearchResult batchResult = new() { Success = false };

            try
            {
                log.WriteLine($"Finding module '{moduleName}'...");
                ModuleInfo moduleInfo = ModuleHelper.GetModuleInfo(processHandle, moduleName) ?? throw new InvalidOperationException($"Module '{moduleName}' not found in process {processId}");
                batchResult.Module = moduleInfo;

                if (outputFormat == OutputFormat.Human)
                {
                    log.WriteLine(moduleInfo.ToString());
                    log.WriteLine();
                }

                string symbolPath = Environment.GetEnvironmentVariable("_NT_SYMBOL_PATH");
                if (string.IsNullOrEmpty(symbolPath))
                {
                    symbolPath = "SRV*C:\\Symbols*https://msdl.microsoft.com/download/symbols";
                }

                log.WriteLine($"Symbol Path: {symbolPath}");
                log.WriteLine();

                SymbolEnumerator enumerator = new(processHandle);

                try
                {
                    log.WriteLine("Initializing symbol handler...");
                    enumerator.InitializeSymbols(symbolPath);

                    log.WriteLine($"Loading symbols for '{moduleInfo.Name}'...");
                    log.WriteLine("(This may take a while if PDB needs to be downloaded)");
                    enumerator.LoadModule(moduleInfo.FullPath, moduleInfo.BaseAddress, moduleInfo.Size);

                    PdbInfo pdbInfo = enumerator.GetPdbInfo();
                    batchResult.PdbInfo = pdbInfo;

                    if (outputFormat == OutputFormat.Human)
                    {
                        log.WriteLine();
                        log.WriteLine(pdbInfo.ToString());
                        log.WriteLine();
                    }

                    log.WriteLine($"Searching for {symbolNames.Count} symbols...");
                    log.WriteLine();

                    // Search for each symbol
                    batchResult.Symbols = new List<SymbolSearchResult>();
                    foreach (string symbolName in symbolNames)
                    {
                        log.WriteLine($"  Looking for '{symbolName}'...");
                        SymbolInfo symbol = enumerator.FindSymbol(symbolName);

                        SymbolSearchResult symbolResult = new()
                        {
                            Success = symbol != null,
                            Symbol = symbol,
                            SearchedSymbolName = symbolName
                        };

                        batchResult.Symbols.Add(symbolResult);
                    }

                    batchResult.Success = true;
                    OutputFormatter.WriteBatchResult(batchResult, outputFormat, Console.Out);
                }
                catch (Exception ex)
                {
                    log.WriteLine(ex.ToString());
                    throw;
                }
                finally
                {
                    enumerator.Cleanup();
                }
            }
            finally
            {
                if (processHandle != IntPtr.Zero)
                {
                    CloseHandle(processHandle);
                }
            }
        }
    }
}

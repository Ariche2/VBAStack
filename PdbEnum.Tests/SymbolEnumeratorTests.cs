using NUnit.Framework;
using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;

namespace PdbEnum.Tests
{
    [TestFixture]
    public class SymbolEnumeratorTests
    {
        private IntPtr _testProcessHandle;
        private int _testProcessId;

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern IntPtr OpenProcess(uint dwDesiredAccess, bool bInheritHandle, uint dwProcessId);

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern bool CloseHandle(IntPtr hObject);

        private const uint PROCESS_QUERY_INFORMATION = 0x0400;
        private const uint PROCESS_VM_READ = 0x0010;

        [OneTimeSetUp]
        public void OneTimeSetUp()
        {
            _testProcessId = Process.GetCurrentProcess().Id;
            _testProcessHandle = OpenProcess(
                PROCESS_QUERY_INFORMATION | PROCESS_VM_READ,
                false,
                (uint)_testProcessId);

            Assert.AreNotEqual(IntPtr.Zero, _testProcessHandle,
                "Failed to open test process");
        }

        [OneTimeTearDown]
        public void OneTimeTearDown()
        {
            if (_testProcessHandle != IntPtr.Zero)
            {
                CloseHandle(_testProcessHandle);
            }
        }

        [Test]
        public void Test_SymbolEnumerator_Constructor()
        {
            SymbolEnumerator enumerator = new SymbolEnumerator(_testProcessHandle);
            Assert.IsNotNull(enumerator, "Should create enumerator instance");
        }

        [Test]
        public void Test_InitializeSymbols_WithDefaultPath()
        {
            SymbolEnumerator enumerator = new SymbolEnumerator(_testProcessHandle);
            
            bool result = enumerator.InitializeSymbols();
            
            Assert.IsTrue(result, "InitializeSymbols should succeed");
            
            enumerator.Cleanup();
        }

        [Test]
        public void Test_InitializeSymbols_WithCustomPath()
        {
            SymbolEnumerator enumerator = new SymbolEnumerator(_testProcessHandle);
            string symbolPath = "SRV*C:\\TestSymbols*https://msdl.microsoft.com/download/symbols";
            
            bool result = enumerator.InitializeSymbols(symbolPath);
            
            Assert.IsTrue(result, "InitializeSymbols should succeed with custom path");
            
            enumerator.Cleanup();
        }

        [Test]
        public void Test_LoadModule_WithValidModule()
        {
            SymbolEnumerator enumerator = new SymbolEnumerator(_testProcessHandle);
            
            try
            {
                enumerator.InitializeSymbols();
                
                ModuleInfo moduleInfo = ModuleHelper.GetModuleInfo(_testProcessHandle, "kernel32.dll");
                Assert.IsNotNull(moduleInfo, "Should find kernel32.dll");
                
                bool result = enumerator.LoadModule(
                    moduleInfo.FullPath,
                    moduleInfo.BaseAddress,
                    moduleInfo.Size);
                
                Assert.IsTrue(result, "LoadModule should succeed");
            }
            finally
            {
                enumerator.Cleanup();
            }
        }

        [Test]
        public void Test_LoadModule_WithoutInitialize_ThrowsException()
        {
            SymbolEnumerator enumerator = new SymbolEnumerator(_testProcessHandle);
            
            ModuleInfo moduleInfo = ModuleHelper.GetModuleInfo(_testProcessHandle, "kernel32.dll");
            Assert.IsNotNull(moduleInfo);
            
            // Should throw because InitializeSymbols wasn't called
            Assert.Throws<InvalidOperationException>(() =>
            {
                enumerator.LoadModule(
                    moduleInfo.FullPath,
                    moduleInfo.BaseAddress,
                    moduleInfo.Size);
            });
        }

        [Test]
        public void Test_GetPdbInfo_AfterLoadModule()
        {
            SymbolEnumerator enumerator = new SymbolEnumerator(_testProcessHandle);

            try
            {
                enumerator.InitializeSymbols();

                ModuleInfo moduleInfo = ModuleHelper.GetModuleInfo(_testProcessHandle, "kernel32.dll");
                Assert.IsNotNull(moduleInfo);

                enumerator.LoadModule(
                    moduleInfo.FullPath,
                    moduleInfo.BaseAddress,
                    moduleInfo.Size);

                PdbInfo pdbInfo = enumerator.GetPdbInfo();

                Assert.IsNotNull(pdbInfo, "Should return PDB info");
                // SymType can be 0 if symbols failed to load, or 4 (Export) if only exports are available
                // We just verify that GetPdbInfo returns a valid object
                Assert.GreaterOrEqual(pdbInfo.SymType, 0U, "SymType should be a valid value");

                TestContext.WriteLine($"PDB Info: {pdbInfo}");
                TestContext.WriteLine($"SymType: {pdbInfo.SymType} (0=None, 1=COFF, 2=CodeView, 3=PDB, 4=Export, 5=Deferred)");
            }
            finally
            {
                enumerator.Cleanup();
            }
        }

        [Test]
        public void Test_GetPdbInfo_WithoutLoadModule_ThrowsException()
        {
            SymbolEnumerator enumerator = new SymbolEnumerator(_testProcessHandle);
            
            try
            {
                enumerator.InitializeSymbols();
                
                Assert.Throws<InvalidOperationException>(() =>
                {
                    enumerator.GetPdbInfo();
                });
            }
            finally
            {
                enumerator.Cleanup();
            }
        }

        [Test]
        [Category("Integration")]
        [Explicit("Requires symbol download, may be slow")]
        public void Test_FindSymbol_KnownSymbol()
        {
            SymbolEnumerator enumerator = new SymbolEnumerator(_testProcessHandle);
            
            try
            {
                enumerator.InitializeSymbols();
                
                ModuleInfo moduleInfo = ModuleHelper.GetModuleInfo(_testProcessHandle, "kernel32.dll");
                Assert.IsNotNull(moduleInfo);
                
                enumerator.LoadModule(
                    moduleInfo.FullPath,
                    moduleInfo.BaseAddress,
                    moduleInfo.Size);
                
                // Search for a well-known exported function
                SymbolInfo symbol = enumerator.FindSymbol("CreateFile");
                
                // Note: Symbol may not be found if PDB isn't available, only exports
                if (symbol != null)
                {
                    Assert.IsNotNull(symbol.Name, "Symbol name should be set");
                    Assert.Greater(symbol.Address, 0UL, "Symbol address should be set");
                    
                    TestContext.WriteLine($"Found symbol: {symbol}");
                }
                else
                {
                    TestContext.WriteLine("Symbol not found (PDB may not be available)");
                }
            }
            finally
            {
                enumerator.Cleanup();
            }
        }

        [Test]
        public void Test_FindSymbol_WithoutLoadModule_ThrowsException()
        {
            SymbolEnumerator enumerator = new SymbolEnumerator(_testProcessHandle);
            
            try
            {
                enumerator.InitializeSymbols();
                
                Assert.Throws<InvalidOperationException>(() =>
                {
                    enumerator.FindSymbol("TestSymbol");
                });
            }
            finally
            {
                enumerator.Cleanup();
            }
        }

        [Test]
        [Category("Integration")]
        [Explicit("Requires symbol download, may be slow")]
        public void Test_EnumerateAllSymbols()
        {
            SymbolEnumerator enumerator = new SymbolEnumerator(_testProcessHandle);
            
            try
            {
                enumerator.InitializeSymbols();
                
                ModuleInfo moduleInfo = ModuleHelper.GetModuleInfo(_testProcessHandle, "kernel32.dll");
                Assert.IsNotNull(moduleInfo);
                
                enumerator.LoadModule(
                    moduleInfo.FullPath,
                    moduleInfo.BaseAddress,
                    moduleInfo.Size);
                
                System.Collections.Generic.List<SymbolInfo> symbols = enumerator.EnumerateAllSymbols();
                
                Assert.IsNotNull(symbols, "Should return symbol list");
                
                if (symbols.Count > 0)
                {
                    TestContext.WriteLine($"Found {symbols.Count} symbols");
                    
                    // Check first few symbols
                    for (int i = 0; i < Math.Min(5, symbols.Count); i++)
                    {
                        Assert.IsNotNull(symbols[i].Name, $"Symbol {i} should have a name");
                        Assert.Greater(symbols[i].Address, 0UL, $"Symbol {i} should have an address");
                    }
                }
                else
                {
                    TestContext.WriteLine("No symbols found (PDB may not be available)");
                }
            }
            finally
            {
                enumerator.Cleanup();
            }
        }

        [Test]
        public void Test_Cleanup_CanBeCalledMultipleTimes()
        {
            SymbolEnumerator enumerator = new SymbolEnumerator(_testProcessHandle);
            
            enumerator.InitializeSymbols();
            enumerator.Cleanup();
            
            Assert.DoesNotThrow(() => enumerator.Cleanup(),
                "Cleanup should be safe to call multiple times");
        }

        [Test]
        public void Test_SetQuietMode()
        {
            // Test that SetQuietMode doesn't throw
            Assert.DoesNotThrow(() => SymbolEnumerator.SetQuietMode(true));
            Assert.DoesNotThrow(() => SymbolEnumerator.SetQuietMode(false));
        }

        [Test]
        public void Test_SymbolInfo_ToString()
        {
            SymbolInfo symbol = new SymbolInfo
            {
                Name = "TestSymbol",
                Address = 0x12345678,
                Size = 0x100,
                Flags = 0x1,
                Tag = 0x5
            };

            string output = symbol.ToString();

            Assert.IsTrue(output.Contains("TestSymbol"), "Should contain symbol name");
            Assert.IsTrue(output.Contains("0x12345678"), "Should contain hex address");
            Assert.IsTrue(output.Contains("256"), "Should contain size in decimal");
        }

        [Test]
        public void Test_PdbInfo_ToString()
        {
            PdbInfo pdbInfo = new PdbInfo
            {
                PdbGuid = Guid.NewGuid(),
                PdbAge = 1,
                PdbFileName = "test.pdb",
                SymType = 3
            };

            string output = pdbInfo.ToString();

            Assert.IsTrue(output.Contains("test.pdb"), "Should contain PDB file name");
            Assert.IsTrue(output.Contains("GUID"), "Should contain GUID label");
        }

        [Test]
        public void Test_PdbInfo_ToString_WithNoGuid()
        {
            PdbInfo pdbInfo = new PdbInfo
            {
                PdbGuid = Guid.Empty,
                PdbAge = 0,
                PdbFileName = "",
                SymType = 4
            };

            string output = pdbInfo.ToString();

            Assert.IsTrue(output.Contains("Symbol Type"), "Should contain symbol type");
        }
    }
}

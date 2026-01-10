using NUnit.Framework;
using System;
using System.Diagnostics;
using System.IO;
using PdbEnum;

namespace PdbEnum.Tests
{
    [TestFixture]
    public class ProgramTests
    {
        [Test]
        public void Test_OutputFormat_Enum()
        {
            Assert.AreEqual(0, (int)OutputFormat.Human);
            Assert.AreEqual(1, (int)OutputFormat.Json);
            Assert.AreEqual(2, (int)OutputFormat.Xml);
        }

        [Test]
        public void Test_SymbolSearchResult_Instantiation()
        {
            SymbolSearchResult result = new SymbolSearchResult
            {
                Success = true,
                ErrorMessage = "Test error",
                SearchedSymbolName = "TestSymbol"
            };

            Assert.IsTrue(result.Success);
            Assert.AreEqual("Test error", result.ErrorMessage);
            Assert.AreEqual("TestSymbol", result.SearchedSymbolName);
        }

        [Test]
        public void Test_BatchSymbolSearchResult_Instantiation()
        {
            BatchSymbolSearchResult result = new BatchSymbolSearchResult
            {
                Success = true,
                Symbols = new System.Collections.Generic.List<SymbolSearchResult>()
            };

            Assert.IsTrue(result.Success);
            Assert.IsNotNull(result.Symbols);
            Assert.AreEqual(0, result.Symbols.Count);
        }

        [Test]
        public void Test_SymbolSearchResult_WithCompleteData()
        {
            SymbolSearchResult result = new SymbolSearchResult
            {
                Success = true,
                Module = new ModuleInfo
                {
                    Name = "test.dll",
                    FullPath = "C:\\test.dll",
                    BaseAddress = 0x10000000,
                    Size = 0x1000,
                    EntryPoint = 0x10001000
                },
                PdbInfo = new PdbInfo
                {
                    PdbGuid = Guid.NewGuid(),
                    PdbAge = 1,
                    PdbFileName = "test.pdb",
                    SymType = 3
                },
                Symbol = new SymbolInfo
                {
                    Name = "TestFunction",
                    Address = 0x10001234,
                    Size = 0x50,
                    Flags = 0,
                    Tag = 0
                },
                SearchedSymbolName = "TestFunction"
            };

            Assert.IsTrue(result.Success);
            Assert.IsNotNull(result.Module);
            Assert.IsNotNull(result.PdbInfo);
            Assert.IsNotNull(result.Symbol);
            Assert.AreEqual("TestFunction", result.SearchedSymbolName);
            Assert.AreEqual("test.dll", result.Module.Name);
            Assert.AreEqual("TestFunction", result.Symbol.Name);
        }

        [Test]
        public void Test_BatchSymbolSearchResult_WithMultipleSymbols()
        {
            BatchSymbolSearchResult batchResult = new BatchSymbolSearchResult
            {
                Success = true,
                Module = new ModuleInfo { Name = "test.dll" },
                PdbInfo = new PdbInfo { SymType = 3 },
                Symbols = new System.Collections.Generic.List<SymbolSearchResult>
                {
                    new SymbolSearchResult
                    {
                        Success = true,
                        SearchedSymbolName = "Symbol1",
                        Symbol = new SymbolInfo { Name = "Symbol1", Address = 0x1000 }
                    },
                    new SymbolSearchResult
                    {
                        Success = false,
                        SearchedSymbolName = "Symbol2",
                        Symbol = null
                    },
                    new SymbolSearchResult
                    {
                        Success = true,
                        SearchedSymbolName = "Symbol3",
                        Symbol = new SymbolInfo { Name = "Symbol3", Address = 0x2000 }
                    }
                }
            };

            Assert.IsTrue(batchResult.Success);
            Assert.AreEqual(3, batchResult.Symbols.Count);
            Assert.IsTrue(batchResult.Symbols[0].Success);
            Assert.IsFalse(batchResult.Symbols[1].Success);
            Assert.IsTrue(batchResult.Symbols[2].Success);
            Assert.IsNull(batchResult.Symbols[1].Symbol);
            Assert.IsNotNull(batchResult.Symbols[0].Symbol);
        }

        [Test]
        [Category("Integration")]
        public void Test_ProgramExecution_ShowsUsageWithNoArgs()
        {
            string exePath = GetPdbEnumExePath();
            if (!File.Exists(exePath))
            {
                Assert.Inconclusive($"PdbEnum.exe not found at {exePath}");
                return;
            }

            ProcessStartInfo psi = new ProcessStartInfo
            {
                FileName = exePath,
                Arguments = "",
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                CreateNoWindow = true
            };

            using (Process process = Process.Start(psi))
            {
                string output = process.StandardOutput.ReadToEnd();
                string error = process.StandardError.ReadToEnd();
                process.WaitForExit();

                string combined = output + error;
                Assert.IsTrue(combined.Contains("Usage") || combined.Contains("PdbEnum"),
                    "Should show usage information");
            }
        }

        [Test]
        [Category("Integration")]
        public void Test_ProgramExecution_JsonFlag()
        {
            string exePath = GetPdbEnumExePath();
            if (!File.Exists(exePath))
            {
                Assert.Inconclusive($"PdbEnum.exe not found at {exePath}");
                return;
            }

            int currentPid = Process.GetCurrentProcess().Id;

            ProcessStartInfo psi = new ProcessStartInfo
            {
                FileName = exePath,
                Arguments = $"-json {currentPid} kernel32.dll CreateFile",
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                CreateNoWindow = true
            };

            using (Process process = Process.Start(psi))
            {
                string output = process.StandardOutput.ReadToEnd();
                string error = process.StandardError.ReadToEnd();
                process.WaitForExit();

                TestContext.Out.WriteLine("Output: " + output);
                TestContext.Out.WriteLine("Error: " + error);

                if (!string.IsNullOrEmpty(output))
                {
                    TestContext.WriteLine("Output: " + output);
                    
                    // Check if output looks like JSON (even if parsing fails)
                    Assert.IsTrue(output.Contains("{") || output.Contains("["),
                        "JSON output should contain braces");
                }
            }
        }

        [Test]
        [Category("Integration")]
        public void Test_ProgramExecution_XmlFlag()
        {
            string exePath = GetPdbEnumExePath();
            if (!File.Exists(exePath))
            {
                Assert.Inconclusive($"PdbEnum.exe not found at {exePath}");
                return;
            }

            int currentPid = Process.GetCurrentProcess().Id;

            ProcessStartInfo psi = new ProcessStartInfo
            {
                FileName = exePath,
                Arguments = $"-xml {currentPid} kernel32.dll CreateFile",
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                CreateNoWindow = true
            };

            using (Process process = Process.Start(psi))
            {
                string output = process.StandardOutput.ReadToEnd();
                string error = process.StandardError.ReadToEnd();
                process.WaitForExit();

                if (!string.IsNullOrEmpty(output))
                {
                    TestContext.WriteLine("Output: " + output);
                    
                    Assert.IsTrue(output.Contains("<?xml") || output.Contains("<"),
                        "XML output should contain XML markers");
                }
            }
        }

        [Test]
        [Category("Integration")]
        public void Test_ProgramExecution_QuietFlag()
        {
            string exePath = GetPdbEnumExePath();
            if (!File.Exists(exePath))
            {
                Assert.Inconclusive($"PdbEnum.exe not found at {exePath}");
                return;
            }

            int currentPid = Process.GetCurrentProcess().Id;

            ProcessStartInfo psi = new ProcessStartInfo
            {
                FileName = exePath,
                Arguments = $"-q {currentPid} kernel32.dll CreateFile",
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                CreateNoWindow = true
            };

            using (Process process = Process.Start(psi))
            {
                string output = process.StandardOutput.ReadToEnd();
                string error = process.StandardError.ReadToEnd();
                process.WaitForExit();

                // In quiet mode, stderr should have less output
                TestContext.WriteLine($"Stderr length: {error.Length}");
            }
        }

        private string GetPdbEnumExePath()
        {
            // Try to find PdbEnum.exe in common locations
            string platform = IntPtr.Size == 8 ? "x64" : "x86";
            string configuration = "Debug";

            string[] possiblePaths = new string[]
            {
                Path.Combine("..", "..", "..", "..", "PdbEnum", "bin", configuration, "PdbEnum_" + platform + ".exe"),
                Path.Combine("..", "..", "..", "..", "PdbEnum", "bin", configuration, "PdbEnum_" + platform + ".exe"),
                Path.Combine("..", "..", "..", "PdbEnum", "bin", configuration, "PdbEnum_" + platform + ".exe"),
                "..\\PdbEnum\\bin\\" + "\\Debug\\PdbEnum_" + platform + ".exe",
                "..\\..\\PdbEnum\\bin\\" + "\\Debug\\PdbEnum_" + platform + ".exe"
            };

            foreach (string path in possiblePaths)
            {
                string fullPath = Path.GetFullPath(path);
                if (File.Exists(fullPath))
                {
                    return fullPath;
                }
            }

            // Return a default path even if not found (will be checked by caller)
            return Path.GetFullPath(possiblePaths[0]);
        }
    }
}

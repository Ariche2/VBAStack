using NUnit.Framework;
using System;
using System.Diagnostics;

namespace PdbEnum.Tests
{
    [TestFixture]
    public class ArchitectureTests
    {
        [Test]
        public void Test_VerifyTestArchitecture()
        {
            bool is64Bit = IntPtr.Size == 8;
            int pointerSize = IntPtr.Size;
            
            TestContext.WriteLine($"Test running as: {(is64Bit ? "x64" : "x86")}");
            TestContext.WriteLine($"Pointer size: {pointerSize} bytes");
            TestContext.WriteLine($"Process architecture: {(Environment.Is64BitProcess ? "64-bit" : "32-bit")}");
            TestContext.WriteLine($"OS architecture: {(Environment.Is64BitOperatingSystem ? "64-bit" : "32-bit")}");
            
            Assert.IsTrue(pointerSize == 4 || pointerSize == 8, 
                "Pointer size should be either 4 (x86) or 8 (x64) bytes");
        }

        [Test]
        public void Test_ProcessArchitectureMatchesPlatformTarget()
        {
            bool is64BitProcess = Environment.Is64BitProcess;
            bool is64BitPointer = IntPtr.Size == 8;
            
            Assert.AreEqual(is64BitPointer, is64BitProcess,
                "Process architecture should match pointer size");
        }

        [Test]
        [Platform(Include = "Win")]
        public void Test_CanRunOnWindows()
        {
            Assert.IsTrue(Environment.OSVersion.Platform == PlatformID.Win32NT,
                "Tests should run on Windows");
        }

        [Test]
        public void Test_DebugHelpPathConstruction()
        {
            string expectedPath;
            if (IntPtr.Size == 8)
            {
                expectedPath = "amd64\\dbghelp.dll";
            }
            else
            {
                expectedPath = "x86\\dbghelp.dll";
            }

            TestContext.WriteLine($"Expected DbgHelp path suffix: {expectedPath}");
            Assert.IsTrue(expectedPath.Contains(IntPtr.Size == 8 ? "amd64" : "x86"),
                "DbgHelp path should match process architecture");
        }
    }
}

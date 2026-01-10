using NUnit.Framework;
using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using PdbEnum;

namespace PdbEnum.Tests
{
    [TestFixture]
    public class ModuleHelperTests
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
            // Use current process for testing
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
        public void Test_GetModuleInfo_ReturnsCurrentProcessModule()
        {
            // Get a known module from current process (kernel32.dll is always loaded)
            ModuleInfo moduleInfo = ModuleHelper.GetModuleInfo(_testProcessHandle, "kernel32.dll");

            Assert.IsNotNull(moduleInfo, "Should find kernel32.dll");
            StringAssert.AreEqualIgnoringCase("kernel32.dll", moduleInfo.Name, "Module name should match (case-insensitive)");
            Assert.IsNotNull(moduleInfo.FullPath, "FullPath should not be null");
            Assert.IsTrue(moduleInfo.FullPath.EndsWith("kernel32.dll", StringComparison.OrdinalIgnoreCase),
                "FullPath should end with kernel32.dll");
            Assert.Greater(moduleInfo.BaseAddress, 0UL, "BaseAddress should be greater than 0");
            Assert.Greater(moduleInfo.Size, 0U, "Size should be greater than 0");
        }

        [Test]
        public void Test_GetModuleInfo_CaseInsensitive()
        {
            // Test case insensitivity
            ModuleInfo lower = ModuleHelper.GetModuleInfo(_testProcessHandle, "kernel32.dll");
            ModuleInfo upper = ModuleHelper.GetModuleInfo(_testProcessHandle, "KERNEL32.DLL");
            ModuleInfo mixed = ModuleHelper.GetModuleInfo(_testProcessHandle, "KeRnEl32.DlL");

            Assert.IsNotNull(lower, "Should find lowercase");
            Assert.IsNotNull(upper, "Should find uppercase");
            Assert.IsNotNull(mixed, "Should find mixed case");

            Assert.AreEqual(lower.BaseAddress, upper.BaseAddress, 
                "All variations should return same base address");
            Assert.AreEqual(lower.BaseAddress, mixed.BaseAddress, 
                "All variations should return same base address");
        }

        [Test]
        public void Test_GetModuleInfo_NonExistentModule_ReturnsNull()
        {
            ModuleInfo moduleInfo = ModuleHelper.GetModuleInfo(_testProcessHandle, "NonExistentModule12345.dll");

            Assert.IsNull(moduleInfo, "Should return null for non-existent module");
        }

        [Test]
        public void Test_GetModuleInfo_NtDll()
        {
            // ntdll.dll is always loaded in every process
            ModuleInfo moduleInfo = ModuleHelper.GetModuleInfo(_testProcessHandle, "ntdll.dll");

            Assert.IsNotNull(moduleInfo, "Should find ntdll.dll");
            StringAssert.AreEqualIgnoringCase("ntdll.dll", moduleInfo.Name, "Module name should match (case-insensitive)");
            Assert.Greater(moduleInfo.BaseAddress, 0UL);
            Assert.Greater(moduleInfo.Size, 0U);
            // Note: EntryPoint may be 0 for some system DLLs, so we don't assert it must be > 0
        }

        [Test]
        public void Test_GetModuleInfo_InvalidProcessHandle_ThrowsException()
        {
            IntPtr invalidHandle = IntPtr.Zero;

            Assert.Throws<InvalidOperationException>(() =>
            {
                ModuleHelper.GetModuleInfo(invalidHandle, "kernel32.dll");
            });
        }

        [Test]
        public void Test_ModuleInfo_ToString_ContainsExpectedInformation()
        {
            ModuleInfo moduleInfo = ModuleHelper.GetModuleInfo(_testProcessHandle, "kernel32.dll");
            Assert.IsNotNull(moduleInfo);

            string output = moduleInfo.ToString();

            Assert.IsTrue(output.IndexOf("kernel32.dll", StringComparison.OrdinalIgnoreCase) >= 0, 
                "ToString should contain module name (case-insensitive)");
            Assert.IsTrue(output.Contains("0x"), "ToString should contain hex addresses");
            Assert.IsTrue(output.Contains("Base Address"), "ToString should label base address");
            Assert.IsTrue(output.Contains("Size"), "ToString should label size");
        }

        [Test]
        public void Test_ModuleInfo_BaseAddress_IsAligned()
        {
            ModuleInfo moduleInfo = ModuleHelper.GetModuleInfo(_testProcessHandle, "kernel32.dll");
            Assert.IsNotNull(moduleInfo);

            // Module base addresses are typically page-aligned (multiple of 4KB or 64KB)
            ulong pageSize = 0x1000; // 4KB
            Assert.AreEqual(0UL, moduleInfo.BaseAddress % pageSize,
                "Base address should be page-aligned");
        }

        [Test]
        public void Test_GetModuleInfo_MultipleModules()
        {
            // Test that we can get info for multiple different modules
            ModuleInfo kernel32 = ModuleHelper.GetModuleInfo(_testProcessHandle, "kernel32.dll");
            ModuleInfo ntdll = ModuleHelper.GetModuleInfo(_testProcessHandle, "ntdll.dll");

            Assert.IsNotNull(kernel32, "Should find kernel32.dll");
            Assert.IsNotNull(ntdll, "Should find ntdll.dll");
            Assert.AreNotEqual(kernel32.BaseAddress, ntdll.BaseAddress,
                "Different modules should have different base addresses");
        }
    }
}

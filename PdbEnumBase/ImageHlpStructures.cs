using System;
using System.Runtime.InteropServices;

namespace PdbEnum
{
    // Reference: https://learn.microsoft.com/en-us/windows/win32/api/dbghelp/ns-dbghelp-imagehlp_module64
    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Ansi)]
    internal struct IMAGEHLP_MODULE64
    {
        public uint SizeOfStruct;
        public ulong BaseOfImage;
        public uint ImageSize;
        public uint TimeDateStamp;
        public uint CheckSum;
        public uint NumSyms;
        public uint SymType; // SYM_TYPE enum

        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 32)]
        public string ModuleName;

        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 256)]
        public string ImageName;

        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 256)]
        public string LoadedImageName;

        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 256)]
        public string LoadedPdbName;

        public uint CVSig;

        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 780)] // MAX_PATH * 3 = 260 * 3
        public string CVData;

        public uint PdbSig;
        public Guid PdbSig70;
        public uint PdbAge;

        [MarshalAs(UnmanagedType.Bool)]
        public bool PdbUnmatched;

        [MarshalAs(UnmanagedType.Bool)]
        public bool DbgUnmatched;

        [MarshalAs(UnmanagedType.Bool)]
        public bool LineNumbers;

        [MarshalAs(UnmanagedType.Bool)]
        public bool GlobalSymbols;

        [MarshalAs(UnmanagedType.Bool)]
        public bool TypeInfo;

        [MarshalAs(UnmanagedType.Bool)]
        public bool SourceIndexed;

        [MarshalAs(UnmanagedType.Bool)]
        public bool Publics;

        public uint MachineType;
        public uint Reserved;
    }

        // Reference: https://learn.microsoft.com/en-us/windows/win32/api/dbghelp/ns-dbghelp-symbol_info
        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Ansi)]
        internal struct SYMBOL_INFO
        {
            public uint SizeOfStruct;
            public uint TypeIndex;
            public ulong Reserved1;
            public ulong Reserved2;
            public uint Index;
            public uint Size;
            public ulong ModBase;
            public uint Flags;
            public ulong Value;
            public ulong Address;
            public uint Register;
            public uint Scope;
            public uint Tag;
            public uint NameLen;
            public uint MaxNameLen;

            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 2000)]
            public string Name;
        }
    }

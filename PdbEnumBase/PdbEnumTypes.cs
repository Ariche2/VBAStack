using System;
using System.Xml.Serialization;
using System.Runtime.Serialization;

namespace PdbEnum
{
    [Serializable]
    [DataContract]
    public class ModuleInfo
    {
        [DataMember]
        public string Name { get; set; }
        [DataMember]
        public string FullPath { get; set; }

        [XmlElement("BaseAddress")]
        [DataMember]
        public ulong BaseAddress { get; set; }

        [DataMember]
        public uint Size { get; set; }

        [XmlElement("EntryPoint")]
        [DataMember]
        public ulong EntryPoint { get; set; }

        public override string ToString()
        {
            return $"Module: {Name}\n  Path: {FullPath}\n  Base Address: 0x{BaseAddress:X}\n  Size: {Size} bytes\n  Entry Point: 0x{EntryPoint:X}";
        }
    }
    [Serializable]
    [DataContract]
    public class SymbolInfo
    {
        [DataMember]
        public string Name { get; set; }

        [XmlElement("Address")]
        [DataMember]
        public ulong Address { get; set; }

        [DataMember]
        public uint Size { get; set; }

        [XmlElement("Flags")]
        [DataMember]
        public uint Flags { get; set; }

        [DataMember]
        public uint Tag { get; set; }

        public override string ToString()
        {
            return $"Symbol: {Name}\n  Address: 0x{Address:X}\n  Size: {Size} bytes\n  Flags: 0x{Flags:X}\n  Tag: {Tag}";
        }
    }

    [Serializable]
    [DataContract]
    public class PdbInfo
    {
        [DataMember]
        public Guid PdbGuid { get; set; }
        [DataMember]
        public uint PdbAge { get; set; }
        [DataMember]
        public string PdbFileName { get; set; }
        [DataMember]
        public uint SymType { get; set; }

        private string GetSymbolTypeName()
        {
            // If SymType claims PDB but we have no valid PDB GUID/filename, correct the type name
            if (SymType == 3 && (PdbGuid == Guid.Empty || string.IsNullOrEmpty(PdbFileName)))
            {
                return "Export";
            }

            switch (SymType)
            {
                case 0: return "None";
                case 1: return "COFF";
                case 2: return "CodeView";
                case 3: return "PDB";
                case 4: return "Export";
                case 5: return "Deferred";
                case 6: return "SYM";
                case 7: return "DIA";
                case 8: return "Virtual";
                default: return $"Unknown ({SymType})";
            }
        }

        public override string ToString()
        {
            string symTypeStr = GetSymbolTypeName();
            if (PdbGuid == Guid.Empty)
            {
                return $"PDB Information:\n  Symbol Type: {symTypeStr}\n  PDB File: {(string.IsNullOrEmpty(PdbFileName) ? "(No PDB loaded - using exports only)" : PdbFileName)}";
            }
            return $"PDB Information:\n  GUID: {PdbGuid:D}\n  Age: {PdbAge}\n  PDB File: {PdbFileName}\n  Symbol Type: {symTypeStr}";
        }
    }
    [Serializable]
    [DataContract]
    public class SymbolSearchResult
    {
        [DataMember]
        public ModuleInfo Module { get; set; }
        [DataMember]
        public PdbInfo PdbInfo { get; set; }
        [DataMember]
        public SymbolInfo Symbol { get; set; }
        [DataMember]
        public bool Success { get; set; }
        [DataMember]
        public string ErrorMessage { get; set; }
        [DataMember]
        public string SearchedSymbolName { get; set; }
    }

    [Serializable]
    [DataContract]
    public class BatchSymbolSearchResult
    {
        [DataMember]
        public ModuleInfo Module { get; set; }
        [DataMember]
        public PdbInfo PdbInfo { get; set; }
        [DataMember]
        public System.Collections.Generic.List<SymbolSearchResult> Symbols { get; set; }
        [DataMember]
        public bool Success { get; set; }
        [DataMember]
        public string ErrorMessage { get; set; }
    }
}

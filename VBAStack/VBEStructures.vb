Imports System.Runtime.InteropServices

''' <summary>
''' Internal VBE structures for direct callstack manipulation.
''' </summary>
Friend Module VBEStructures

    ''' <summary>
    ''' Execution frame structure - forms a linked list of active VBA calls.
    ''' Reverse engineered from VBE7.dll.
    ''' Size: 0x4C (76 bytes) on x86
    ''' </summary>
    <StructLayout(LayoutKind.Sequential)>
    Friend Structure EXFRAME
        ''' <summary>Offset 0x00: Pointer to next EXFRAME in linked list (previous call in callstack)</summary>
        Public lpNext As Integer

        ''' <summary>Offset 0x04 (x86): Unknown field, referenced in ExecFillNtexInfoFromExframe as field_0x4</summary>
        Public field_0x4 As Integer

        ''' <summary>Offset 0x08 (x86): Unknown field</summary>
        Public field_0x8 As Integer

        ''' <summary>Offset 0x0C (x86): Pointer to runtime member info (RTMI) - critical for resolving function names.
        ''' Present at 0x18 on x64, confirming last 2 fields are definitely pointers.</summary>
        Public lpRTMI As Integer

        ''' <summary>Offset 0x10 (x86): Unknown field</summary>
        Public field_0x10 As Integer

        ''' <summary>Offset 0x14 (x86): Unknown field</summary>
        Public field_0x14 As Integer

        ''' <summary>Offset 0x18 (x86): Unknown field</summary>
        Public field_0x18 As Integer

        ''' <summary>Offset 0x1C (x86): Optional value (integer) - used in ExecGetNtexCountFromExframe, checked for 0 or -1</summary>
        Public optionalValue As Integer

        ''' <summary>Offset 0x20 (x86): Optional pointer 1 - checked for null in ExecGetNtexCountFromExframe</summary>
        Public lpOptional1 As Integer

        ''' <summary>Offset 0x24 (x86): Unknown field</summary>
        Public field_0x24 As Integer

        ''' <summary>Offset 0x28 (x86): Count of local variables in this frame</summary>
        Public cLocalVars As Integer

        ''' <summary>Offset 0x2C (x86): Optional pointer 2 - checked for null in ExecGetNtexCountFromExframe</summary>
        Public lpOptional2 As Integer

        ''' <summary>Offset 0x30 (x86): Unknown field</summary>
        Public field_0x30 As Integer

        ''' <summary>Offset 0x34 (x86): Unknown field</summary>
        Public field_0x34 As Integer

        ''' <summary>Offset 0x38 (x86): Unknown field</summary>
        Public field_0x38 As Integer

        ''' <summary>Offset 0x3C (x86): Unknown field</summary>
        Public field_0x3C As Integer

        ''' <summary>Offset 0x40 (x86): Unknown field</summary>
        Public field_0x40 As Integer

        ''' <summary>Offset 0x44 (x86): Unknown field</summary>
        Public field_0x44 As Integer

        ''' <summary>Offset 0x48 (x86): Current instruction pointer (ptex) - points to current p-code instruction being executed</summary>
        Public lpCurrentIP As Integer

        ' Note: Local variables are stored at negative offsets from the EXFRAME base address
        ' They start at (EXFRAME_address - 0x28 - RTMI->cbStackFrame) and go downward
        ' Each local var is 4 bytes and addressed at (base - 4*index)
    End Structure

    ''' <summary>
    ''' RTMI (Runtime Member Info) structure - describes a VBA procedure at runtime.
    ''' Reverse engineered from VBE7.dll.
    ''' Size: Variable (at least 0x18+ bytes, contains pointers to additional structures)
    ''' </summary>
    <StructLayout(LayoutKind.Sequential)>
    Friend Structure RTMI
        ''' <summary>Offset 0x00: Pointer to ObjectInfo structure (link back to parent object)</summary>
        Public lpObjectInfo As Integer

        ''' <summary>Offset 0x02 (x86): Module index (ushort) - used to look up ExecMod via ExecProj::Pexecmod</summary>
        Public moduleIndex As UShort

        ''' <summary>Offset 0x04 (x86): Pointer to structure containing ExecProj at offset +0x04</summary>
        Public lpExecProj As Integer

        ''' <summary>Offset 0x06 (x86): Stack frame size (ushort) - cbStackFrame
        ''' Used by ExecFillNtexInfoFromExframe to calculate local variable addresses
        ''' Local vars start at: (EXFRAME_address - 0x28 - cbStackFrame) and go downward</summary>
        Public cbStackFrame As UShort

        ''' <summary>Offset 0x08 (x86): Unknown field</summary>
        Public field_0x8 As Integer

        ''' <summary>Offset 0x0C (x86): Unknown field</summary>
        Public field_0xC As Integer

        ''' <summary>Offset 0x10 (x86): Unknown field</summary>
        Public field_0x10 As Integer

        ''' <summary>Offset 0x14 (x86): Unknown field</summary>
        Public field_0x14 As Integer

        ''' <summary>Offset 0x18 (x86): Pointer to RESDESCTBL structure (resource descriptor table)
        ''' Used in SerReadRTMI for reading/writing serialized data</summary>
        Public lpResDescTbl As Integer

        ' Note: Additional fields may exist beyond 0x18, but not yet mapped
    End Structure

    ''' <summary>
    ''' ObjectInfo structure - defines an Object and provides information to its methods and constants.
    ''' See Notes_on_VB_Structures.txt.
    ''' Size: 0x38 bytes
    ''' </summary>
    <StructLayout(LayoutKind.Sequential)>
    Friend Structure OBJECTINFO
        ''' <summary>Offset 0x00: Always 1 after compilation</summary>
        Public wRefCount As UShort

        ''' <summary>Offset 0x02: Index of this Object</summary>
        Public wObjectIndex As UShort

        ''' <summary>Offset 0x04: Pointer to the Object Table</summary>
        Public lpObjectTable As Integer

        ''' <summary>Offset 0x08: Zero after compilation. Used in IDE only</summary>
        Public lpIdeData As Integer

        ''' <summary>Offset 0x0C: Pointer to Private Object Descriptor</summary>
        Public lpPrivateObject As Integer

        ''' <summary>Offset 0x10: Always -1 after compilation</summary>
        Public dwReserved As Integer

        ''' <summary>Offset 0x14: Unused</summary>
        Public dwNull As Integer

        ''' <summary>Offset 0x18: Back-Pointer to Public Object Descriptor</summary>
        Public lpObject As Integer

        ''' <summary>Offset 0x1C: Pointer to in-memory Project Object</summary>
        Public lpProjectData As Integer

        ''' <summary>Offset 0x20: Number of Methods</summary>
        Public wMethodCount As UShort

        ''' <summary>Offset 0x22: Zeroed out after compilation. IDE only</summary>
        Public wMethodCount2 As UShort

        ''' <summary>Offset 0x24: Pointer to Array of RTMI pointers (one per method)</summary>
        Public lpMethods As Integer

        ''' <summary>Offset 0x28: Number of Constants in Constant Pool</summary>
        Public wConstants As UShort

        ''' <summary>Offset 0x2A: Constants to allocate in Constant Pool</summary>
        Public wMaxConstants As UShort

        ''' <summary>Offset 0x2C: Valid in IDE only</summary>
        Public lpIdeData2 As Integer

        ''' <summary>Offset 0x30: Valid in IDE only</summary>
        Public lpIdeData3 As Integer

        ''' <summary>Offset 0x34: Pointer to Constants Pool</summary>
        Public lpConstants As Integer
    End Structure

    ''' <summary>
    ''' ObjectTable structure - contains pointers to objects and project data.
    ''' See Notes_on_VB_Structures.txt.
    ''' Size: 0x54 bytes
    ''' </summary>
    <StructLayout(LayoutKind.Sequential)>
    Friend Structure OBJECTTABLE
        ''' <summary>Offset 0x00: Unused after compilation, always 0</summary>
        Public lpHeapLink As IntPtr

        ''' <summary>Offset 0x04: Pointer to VB Project Exec COM Object (ExecProj)</summary>
        Public lpExecProj As IntPtr

        ''' <summary>Offset 0x08: Secondary Project Information</summary>
        Public lpProjectInfo2 As IntPtr

        ''' <summary>Offset 0x0C: Always set to -1 after compiling. Unused</summary>
        Public dwReserved As IntPtr

        ''' <summary>Offset 0x10: Not used in compiled mode</summary>
        Public dwNull As IntPtr

        ''' <summary>Offset 0x14: Pointer to in-memory Project Data</summary>
        Public lpProjectObject As IntPtr

        ''' <summary>Offset 0x18: GUID of the Object Table (16 bytes)</summary>
        Public guid As Guid

        ''' <summary>Offset 0x28: Internal flag used during compilation</summary>
        Public fCompileState As UShort

        ''' <summary>Offset 0x2A: Total objects present in Project</summary>
        Public dwTotalObjects As UShort

        ''' <summary>Offset 0x2C: Equal to above after compiling</summary>
        Public dwCompiledObjects As UShort

        ''' <summary>Offset 0x2E: Usually equal to above after compile</summary>
        Public dwObjectsInUse As UShort

        ''' <summary>Offset 0x30: Pointer to Object Descriptors</summary>
        Public lpObjectArray As IntPtr

        ''' <summary>Offset 0x34: Flag/Pointer used in IDE only</summary>
        Public fIdeFlag As IntPtr

        ''' <summary>Offset 0x38: Flag/Pointer used in IDE only</summary>
        Public lpIdeData As IntPtr

        ''' <summary>Offset 0x3C: Flag/Pointer used in IDE only</summary>
        Public lpIdeData2 As IntPtr

        ''' <summary>Offset 0x40: Pointer to Project Name (ANSI string)</summary>
        Public lpszProjectName As IntPtr

        ''' <summary>Offset 0x44: LCID of Project</summary>
        Public dwLcid As Integer

        ''' <summary>Offset 0x48: Alternate LCID of Project</summary>
        Public dwLcid2 As Integer

        ''' <summary>Offset 0x4C: Flag/Pointer used in IDE only</summary>
        Public lpIdeData3 As Integer

        ''' <summary>Offset 0x50: Template Version of Structure</summary>
        Public dwIdentifier As Integer
    End Structure

    ''' <summary>
    ''' Public Object Descriptor - describes a VBA module/class and its methods.
    ''' See Notes_on_VB_Structures.txt.
    ''' Size: 0x30 bytes
    ''' </summary>
    <StructLayout(LayoutKind.Sequential)>
    Friend Structure PUBLIC_OBJECT_DESCRIPTOR
        ''' <summary>Offset 0x00: Pointer to the Object Info for this Object</summary>
        Public lpObjectInfo As Integer

        ''' <summary>Offset 0x04: Always set to -1 after compiling</summary>
        Public dwReserved As Integer

        ''' <summary>Offset 0x08: Pointer to Public Variable Size integers</summary>
        Public lpPublicBytes As Integer

        ''' <summary>Offset 0x0C: Pointer to Static Variable Size integers</summary>
        Public lpStaticBytes As Integer

        ''' <summary>Offset 0x10: Pointer to Public Variables in DATA section</summary>
        Public lpModulePublic As Integer

        ''' <summary>Offset 0x14: Pointer to Static Variables in DATA section</summary>
        Public lpModuleStatic As Integer

        ''' <summary>Offset 0x18: Name of the Object/Module (ANSI string pointer)</summary>
        Public lpszObjectName As Integer

        ''' <summary>Offset 0x1C: Number of Methods in Object</summary>
        Public dwMethodCount As Integer

        ''' <summary>Offset 0x20: Pointer to array of method name string pointers (ANSI)</summary>
        Public lpMethodNames As Integer

        ''' <summary>Offset 0x24: Offset to where to copy Static Variables</summary>
        Public bStaticVars As Integer

        ''' <summary>Offset 0x28: Flags defining the Object Type</summary>
        Public fObjectType As Integer

        ''' <summary>Offset 0x2C: Not valid after compilation</summary>
        Public dwNull As Integer
    End Structure

    ''' <summary>
    ''' Callstack frame wrapper (only used in MDB files with debug info)
    ''' This structure exists at g_pUIApp + 0x2f0 in MDB files but not in MDEs
    ''' </summary>
    <StructLayout(LayoutKind.Sequential)>
    Friend Structure CallstackFrame
        ''' <summary>Pointer to next frame in linked list</summary>
        Public lpNext As Integer

        ''' <summary>Frame type: 1=function, 2=?, 3=?, 5=?</summary>
        Public frameType As Integer

        ''' <summary>Pointer to EXFRAME structure</summary>
        Public lpExFrame As Integer
    End Structure

End Module

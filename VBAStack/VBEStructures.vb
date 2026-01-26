Imports System.Runtime.InteropServices

''' <summary>
''' Internal VBE structures for direct callstack manipulation.
''' 
''' PROJECT STRUCTURE (Reverse Engineered via Ghidra):
''' ===================================================
''' The Project structure is the base class for VBA projects in VBE7.dll.
''' Size: ~256 bytes (exact size varies, but important fields are known)
''' 
''' Key Offsets:
''' +0x00  vtable         - Virtual function table pointer
''' +0x04  pNext          - Next project in linked list
''' +0x08  pPrev          - Previous project in linked list
''' +0xA4  pITypeLib      - Pointer to ITypeLib interface object
''' 
''' The pITypeLib at offset 0xA4 is critical for getting project information:
''' - It points to an object that implements ITypeLib COM interface
''' - ITypeLib vtable offset +0x24 (36 decimal) is GetDocumentation method
''' - GetDocumentation retrieves type library metadata (not code comments)
''' - When called with index=-1 (0xFFFFFFFF), it returns the library name = project name
''' - Function signature: HRESULT GetDocumentation(int index, BSTR* pName, BSTR* pDocString, DWORD* pdwHelpContext, BSTR* pHelpFile)
''' - VBE uses: GetDocumentation(-1, &pName, NULL, NULL, NULL) to get just the project name
''' 
''' GEN_PROJECT STRUCTURE (Derived from Project):
''' =============================================
''' GEN_PROJECT is a larger, more complex structure that inherits from Project.
''' Size: 3320 bytes (0xCF8)
''' 
''' Key Offsets:
''' +0x00  vtable_ITypeLib         - ITypeLib vtable
''' +0x04  vtable_ICreateTypeLib   - ICreateTypeLib vtable  
''' +0x08  vtable_IVbaProject      - IVbaProject vtable
''' +0x0C  pPrev                   - Previous GEN_PROJECT in global linked list
''' +0x10  pNext                   - Next GEN_PROJECT in global linked list
''' +0x24  pInternalData           - Pointer to internal data at offset +0x28
''' +0x28  internalData[144]       - Internal data buffer
''' +0xB8  FILTITER structure
''' +0xD4  BLKDESC32 structure
''' +0xE0  BLK_DESC structure
''' +0xF4  BLKMGR32 structure
''' +0x190 VERMGR structure
''' +0x1A4 GENPROJ_TYPEBIND structure
''' +0x214 NAMMGR structure (large: 2604 bytes)
''' +0xC40 BLKDESC32 structure
''' +0xC88 BLKDESC32 structure
''' +0xCC8 GUIDMGR structure
''' 
''' Global Variables:
''' DAT_1025d540 - Pointer to first GEN_PROJECT in linked list
''' DAT_1025d544 - Pointer to last GEN_PROJECT in linked list
''' 
''' Key Functions:
''' - GetProjectName (0x10035d36): Retrieves project name using pITypeLib->GetDocumentation
''' - GEN_PROJECT constructor (0x100cd4bf): Initializes the structure and sub-objects
''' - GetCurrentGenProj (0x100b4b3d): Gets the currently active GEN_PROJECT
''' - _TipGetProjName@8 (0x100ad441): Helper that calls ITypeLib::GetDocumentation
''' 
''' EXECPROJ STRUCTURE (Reverse Engineered via Ghidra):
''' ===================================================
''' ExecProj is the execution-time representation of a VBA project.
''' Size: At least 0x1028 bytes (based on field offsets)
''' 
''' Key Offsets:
''' +0x1024  pModuleArray   - Pointer to array of ExecMod* pointers (indexed by module index)
''' 
''' Key Functions:
''' - ExecProj::Pexecmod(ushort moduleIndex) @ 0x100af00b: Returns ExecMod* from module index
'''   Implementation: return pModuleArray ? pModuleArray[moduleIndex] : NULL
''' - BilsymGetExecProj(FEDSYM*) @ 0x100f4f50: Gets ExecProj from a symbol
''' 
''' EXECMOD STRUCTURE (Reverse Engineered via Ghidra):
''' ==================================================
''' ExecMod is the execution-time representation of a VBA module.
''' Size: At least 0x18 bytes (based on field offsets)
''' 
''' Key Offsets:
''' +0x00  pVTable        - Virtual function table pointer
''' +0x14  pTypeInfo      - Pointer to BASIC_TYPEINFO structure
''' 
''' Key Functions:
''' - ExecMod constructor @ 0x101a0265: Initializes the structure
''' 
''' BASIC_TYPEINFO STRUCTURE (Reverse Engineered via Ghidra):
''' =========================================================
''' Contains type information for a module.
''' Size: At least 0x18 bytes (based on field offsets)
''' 
''' Key Offsets:
''' +0x14  pTypeRoot      - Pointer to BASIC_TYPEROOT structure
''' 
''' RTMI (Runtime Member Info) STRUCTURE (Reverse Engineered via Ghidra):
''' =====================================================================
''' Runtime member information structure - describes a VBA procedure at runtime.
''' Size: Variable (contains additional data structures)
''' 
''' Key Offsets:
''' +0x00  pObjectInfo    - Pointer to ObjectInfo structure (link back to parent object)
''' +0x02  moduleIndex    - Module index (ushort) - used to look up ExecMod
''' +0x04  pSomething2    - Unknown pointer (dereferenced twice to get ExecProj)
''' +0x06  cbStackFrame   - Stack frame size (ushort) - used to calculate local variable addresses
''' +0x0A  something      - Unknown field
''' +0x18  pResDescTbl    - Pointer to RESDESCTBL structure (resource descriptor table)
''' 
''' OBJECTINFO STRUCTURE (See Notes_on_VB_Structures.txt):
''' ====================================================================
''' The Object Information structure defines an Object and provides various information to its methods.
''' Size: At least 0x38 bytes (based on documented offsets)
''' 
''' Key Offsets
''' +0x00  wRefCount          - Always 1 after compilation
''' +0x02  wObjectIndex       - Index of this Object
''' +0x04  lpObjectTable      - Pointer to the Object Table
''' +0x08  lpIdeData          - Zero after compilation. Used in IDE only
''' +0x0C  lpPrivateObject    - Pointer to Private Object Descriptor
''' +0x10  dwReserved         - Always -1 after compilation
''' +0x14  dwNull             - Unused
''' +0x18  lpObject           - Back-Pointer to Public Object Descriptor
''' +0x1C  lpProjectData      - Pointer to in-memory Project Object
''' +0x20  wMethodCount       - Number of Methods
''' +0x22  wMethodCount2      - Zeroed out after compilation. IDE only
''' +0x24  lpMethods          - Pointer to Array of RTMI pointers (one per method)
''' +0x28  wConstants         - Number of Constants in Constant Pool
''' +0x2A  wMaxConstants      - Constants to allocate in Constant Pool
''' +0x2C  lpIdeData2         - Valid in IDE only
''' +0x30  lpIdeData3         - Valid in IDE only
''' +0x34  lpConstants        - Pointer to Constants Pool
''' 
''' OBJECTTABLE STRUCTURE (See Notes_on_VB_Structures.txt):
''' ========================================================
''' The Object Table structure is pointed by the Project Info Structure.
''' Size: 0x54 bytes
''' 
''' Key Offsets:
''' +0x00  lpHeapLink         - Unused after compilation, always 0
''' +0x04  lpExecProj         - Pointer to VB Project Exec COM Object
''' +0x08  lpProjectInfo2     - Secondary Project Information
''' +0x0C  dwReserved         - Always set to -1 after compiling. Unused
''' +0x10  dwNull             - Not used in compiled mode
''' +0x14  lpProjectObject    - Pointer to in-memory Project Data
''' +0x18  uuidObject         - GUID of the Object Table
''' +0x28  fCompileState      - Internal flag used during compilation
''' +0x2A  dwTotalObjects     - Total objects present in Project
''' +0x2C  dwCompiledObjects  - Equal to above after compiling
''' +0x2E  dwObjectsInUse     - Usually equal to above after compile
''' +0x30  lpObjectArray      - Pointer to Object Descriptors
''' +0x34  fIdeFlag           - Flag/Pointer used in IDE only
''' +0x38  lpIdeData          - Flag/Pointer used in IDE only
''' +0x3C  lpIdeData2         - Flag/Pointer used in IDE only
''' +0x40  lpszProjectName    - Pointer to Project Name (ANSI string)
''' +0x44  dwLcid             - LCID of Project
''' +0x48  dwLcid2            - Alternate LCID of Project
''' +0x4C  lpIdeData3         - Flag/Pointer used in IDE only
''' +0x50  dwIdentifier       - Template Version of Structure
''' 
''' PUBLIC OBJECT DESCRIPTOR (See Notes_on_VB_Structures.txt):
''' =========================================================================
''' The Public Object Descriptor Table is pointed by the Array lpObjectArray in the Object Table.
''' Each Object in the project will have its own. Used by VB for a variety of runtime tasks.
''' Size: 0x30 bytes
''' 
''' Key Offsets:
''' +0x00  lpObjectInfo       - Pointer to the Object Info for this Object
''' +0x04  dwReserved         - Always set to -1 after compiling
''' +0x08  lpPublicBytes      - Pointer to Public Variable Size integers
''' +0x0C  lpStaticBytes      - Pointer to Static Variable Size integers
''' +0x10  lpModulePublic     - Pointer to Public Variables in DATA section
''' +0x14  lpModuleStatic     - Pointer to Static Variables in DATA section
''' +0x18  lpszObjectName     - Name of the Object/Module (ANSI string pointer)
''' +0x1C  dwMethodCount      - Number of Methods in Object
''' +0x20  lpMethodNames      - Pointer to array of method name string pointers (ANSI)
''' +0x24  bStaticVars        - Offset to where to copy Static Variables
''' +0x28  fObjectType        - Flags defining the Object Type
''' +0x2C  dwNull             - Not valid after compilation
''' 
''' Relationship Flow (from EXFRAME to module/function names in compiled code):
''' 1. EXFRAME.pRTMI (+0x0C) points to RTMI structure
''' 2. RTMI.pObjectInfo (+0x00) points to ObjectInfo structure
''' 3. ObjectInfo.lpObjectTable (+0x04) points to ObjectTable structure
''' 4. ObjectTable.lpszProjectName (+0x40) contains the project name string
''' 5. ObjectInfo.lpObject (+0x18) points to Public Object Descriptor (back-pointer)
''' 6. Public Object Descriptor.lpszObjectName (+0x18) contains the module/class name
''' 7. ObjectInfo.lpMethods (+0x24) points to array of RTMI pointers
''' 8. Public Object Descriptor.dwMethodCount (+0x1C) contains the number of methods
''' 9. Search through lpMethods array to find index of matching RTMI pointer
''' 10. Public Object Descriptor.lpMethodNames (+0x20) points to array of method name pointers
''' 11. Use found index to retrieve function name from lpMethodNames array
''' 
''' This flow works in both compiled (MDE) and uncompiled (MDB) VBA projects.
''' 
''' Key Functions Using RTMI:
''' - BtrootOfExframe @ 0x100b608c: Extracts BASIC_TYPEROOT from EXFRAME
'''   Flow: EXFRAME->pRTMI->pSomething2->+0x4->+0x4->Pexecmod(moduleIndex)->pTypeInfo->+0x14->pTypeRoot
''' - GetBtinfoOfExframe @ 0x100b5044: Extracts BASIC_TYPEINFO from EXFRAME
'''   Flow: EXFRAME->pRTMI->pSomething2->+0x4->+0x4->Pexecmod(moduleIndex)->pTypeInfo
''' - epiModule::MemidOfPrtmi @ 0x101aa58c: Gets member ID from RTMI pointer
''' - ExecFillNtexInfoFromExframe @ 0x10137413: Uses RTMI->cbStackFrame to calculate local var addresses
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
    ''' ExecProj structure - execution-time representation of a VBA project.
    ''' Reverse engineered from VBE7.dll.
    ''' Size: At least 0x1028 bytes (based on pModuleArray offset)
    ''' </summary>
    <StructLayout(LayoutKind.Sequential)>
    Friend Structure EXECPROJ
        ' First 0x1024 bytes - internal fields not yet fully mapped
        Private _padding As Integer ' Placeholder for unmapped fields

        ' Only the known critical field is mapped:
        ''' <summary>Offset 0x1024 (x86): Pointer to array of ExecMod* pointers
        ''' Access via ExecProj::Pexecmod(ushort moduleIndex) which returns pModuleArray[moduleIndex]</summary>
        Public lpModuleArray As Integer

        ' Note: This structure is incomplete - only showing the most important field
        ' Use ExecProj::Pexecmod function to safely access module array
    End Structure

    ''' <summary>
    ''' ExecMod structure - execution-time representation of a VBA module.
    ''' Reverse engineered from VBE7.dll.
    ''' Size: At least 0x18 bytes (based on pTypeInfo offset)
    ''' </summary>
    <StructLayout(LayoutKind.Sequential)>
    Friend Structure EXECMOD
        ''' <summary>Offset 0x00: Virtual function table pointer</summary>
        Public lpVTable As Integer

        ''' <summary>Offset 0x04-0x10: Unknown fields (not yet mapped)</summary>
        Private field_0x4 As Integer
        Private field_0x8 As Integer
        Private field_0xC As Integer
        Private field_0x10 As Integer

        ''' <summary>Offset 0x14 (x86): Pointer to BASIC_TYPEINFO structure
        ''' Used to navigate to BASIC_TYPEROOT via BASIC_TYPEINFO->pTypeRoot (+0x14)</summary>
        Public lpTypeInfo As Integer
    End Structure

    ''' <summary>
    ''' BASIC_TYPEINFO structure - contains type information for a module.
    ''' Reverse engineered from VBE7.dll.
    ''' Size: At least 0x18 bytes (based on pTypeRoot offset)
    ''' </summary>
    <StructLayout(LayoutKind.Sequential)>
    Friend Structure BASIC_TYPEINFO
        ''' <summary>Offset 0x00-0x10: Unknown fields (not yet mapped)</summary>
        Private field_0x0 As Integer
        Private field_0x4 As Integer
        Private field_0x8 As Integer
        Private field_0xC As Integer
        Private field_0x10 As Integer

        ''' <summary>Offset 0x14 (x86): Pointer to BASIC_TYPEROOT structure
        ''' BASIC_TYPEROOT is used with GetFrameNames to retrieve module and function names</summary>
        Public lpTypeRoot As Integer
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

    ''' <summary>
    ''' Frame type enumeration
    ''' </summary>
    Friend Enum CallstackFrameType As Integer
        ''' <summary>Regular VBA function/sub</summary>
        RegularFunction = 1
        ''' <summary>Unknown frame type 2</summary>
        Unknown2 = 2
        ''' <summary>Unknown frame type 3</summary>
        Unknown3 = 3
        ''' <summary>Unknown frame type 5</summary>
        Unknown5 = 5
    End Enum

End Module

# VBAStack

A library for retrieving VBA callstack information at runtime from Office applications. This enables debugging and error reporting capabilities for VBA add-ins and COM add-ins targeting Microsoft Office.

Credit to "The Trick" - without finding his [VbTrickTimer](https://github.com/thetrik/VbTrickTimer) code, I would've just accepted that VBE7 doesn't export the functions to do this, and would never have thought to go digging around in VBE7.dll directly.


# NEW: VBAStack directly in VBA!
Managed to get it working directly in VBA, no .NET or COM or anything. So far I've tested in x86 Access 2003, x86 Access 2013, x86 Access 365, x64 Access 2013, and x64 Access 365, and it works across the board.

Just download the [VBAStack.bas](VBAStack.bas) file and import it into your VBA project - notes on usage are included in the file.


## Note on AI use
The vast majority of this code was written with my own two hands, but I will admit to prettifying things (mostly documentation, and the bones of this readme) with AI.

I did discover that it is *absolutely terrible* at debugging stuff when it isn't well-covered ground, though.

## Overview

VBAStack allows you to programmatically retrieve the current VBA call stack from a running VBA application (Excel, Word, Access, etc.).

The library provides **two methods** for retrieving callstacks:

1. **VBEDirectCallstackReader** (Recommended) - Directly walks VBE internal structures by navigating the EXFRAME linked list. Works without requiring PDB symbols or internet access, and supports both compiled (MDE) and uncompiled (MDB) VBA projects.

2. **VBECallstackProvider** (Legacy) - Uses PDB symbols to resolve VBE7.dll function addresses and calls internal VBE functions. Requires internet access for initial symbol download. **Note: This method is marked as obsolete.**

Both methods work by accessing internal undocumented structures and functions inside VBE7.dll. The library needs to be loaded into the same process as the VBE, so it's best deployed as part of a COM add-in or VSTO add-in and called from VBA via the add-in.

I am personally using it with a VSTO addin for MS Access, which makes deployment incredibly easy (not quite as easy as a certain other tool that can get the callstack, but I'm working on it).

***Due to the undocumented nature of these internal VBE structures and functions, this library is provided as-is without any guarantees. Use at your own risk.***

## Requirements

- **VBE 7.0+** (Office 2010 and later)
- **.NET Framework 4.8**

## Architecture

The solution consists of several interconnected projects:

### Core Projects

#### VBAStack (VB.NET)
The main library that provides the high-level API for retrieving VBA callstacks.

**Key Components:**
- `VBEDirectCallstackReader` - **[Recommended]** Directly walks VBE internal structures (EXFRAME linked list, ObjectInfo, ObjectTable) to extract callstack information. Works without PDB symbols.
- `VBECallstackProvider` - **[Obsolete]** Legacy API for getting callstack information via PDB-resolved function pointers
- `VBESymbolResolver` - Communicates with PdbEnum to resolve function addresses from PDB symbols (used by legacy method)
- `VBENativeWrapper` - Marshals calls to native VBE7.DLL functions (used by legacy method)
- `VBEWindowHook` - Manages VBE window visibility during callstack capture (used by legacy method)
- `VBEStructures` - Defines VBE internal data structures for direct stack walking
- `VBAStackLogger` - Provides logging capabilities for debugging

#### PdbEnumBase (C#)
Base library for enumerating symbols from PDB (Program Database) files. Used by the legacy callstack retrieval method. Supports JSON, XML, and human-readable output formats.

**Key Components:**
- `SymbolEnumerator` - Interfaces with DbgHelp.dll to load and enumerate symbols
- `ModuleHelper` - Manages module loading and symbol resolution
- `OutputFormatter` - Formats symbol information for consumption in multiple formats
- `ImageHlpStructures` - P/Invoke structures for DbgHelp API

**Note:** Not recommended for use outside of this project, as it hasn't been tested beyond getting the necessary symbols from VBE7.dll.

#### PdbEnum_x64 / PdbEnum_x86 (C#)
Platform-specific console executables that extract symbol addresses from VBE7.DLL's PDB files. Used only by the legacy VBECallstackProvider method.

These are separate processes to isolate DbgHelp.dll usage, which can interfere with debugging. They download symbols from Microsoft Symbol Servers on first use (~4MB download).

#### NativePtrCaller (C#)
Provides wrappers for calling native function pointers from managed code using C# 9.0 function pointers. Used by the legacy method.

**Key Functions:**
- `EbMode()` - Gets current VBE execution mode
- `EbSetMode()` - Sets VBE execution mode (Design/Break/Run)
- `EbGetCallstackCount()` - Gets number of stack frames
- `ErrGetCallstackString()` - Retrieves formatted callstack string for a frame

## Usage

### Basic Example

VBA error handler in your Access, Excel, Word, etc. VBA module
```vba
Public Sub ExampleVbaProcedure()
    On Error GoTo ErrorHandler
    '... your code here ...
ErrorHandler:
    Dim callstack As String
    ' Recommended: Use the direct method (no PDB symbols needed)
    callstack = Application.COMAddins("MyAddin").Object.GetVbaCallstackDirect()
    
    ' Legacy: Use PDB-based method (requires internet for first-time symbol download)
    'callstack = Application.COMAddins("MyAddin").Object.GetVbaCallstackLegacy(Application.VBE)
    
    MsgBox "An error occurred. Callstack:" & vbCrLf & callstack
    Exit Sub
End Sub
```

VB.Net - within a VSTO or COM add-in
```vb.net
Imports VBAStack

' Recommended: Direct method (no PDB symbols needed, works with MDE files)
Public Function GetVbaCallstackDirect() As String
    Return VBEDirectCallstackReader.GetCallstackString()
End Function

' Legacy: PDB-based method (requires internet access, marked obsolete)
Public Function GetVbaCallstackLegacy(vbe As Object) As String
    Return VBECallstackProvider.GetCallstack(vbe)
End Function
```

C# - same as above, but in C#
```csharp
using VBAStack;

// Recommended: Direct method (no PDB symbols needed, works with MDE files)
public string GetVbaCallstackDirect()
{
    return VBEDirectCallstackReader.GetCallstackString();
}

// Legacy: PDB-based method (requires internet access, marked obsolete)
public string GetVbaCallstackLegacy(object vbe)
{
    return VBECallstackProvider.GetCallstack(vbe);
}
```

### Example Output

```
Module1::ProcessData
Module2::CalculateResults
ThisWorkbook::Workbook_Open
```


## How It Works

### Direct Method (Recommended - VBEDirectCallstackReader)

1. Calls `rtcErrObj()` from VBE7.dll to get a pointer to the VBA error object
2. From the error object, navigates to the global `g_ebThread` variable (at offset 0x18 in the error object)
3. Locates the `g_ExFrameTOS` (top-of-stack) global variable by reading ahead 3 pointer-sizes from `g_ebThread` (these globals are always in the same order)
4. Walks the EXFRAME linked list starting from the top-of-stack pointer
5. For each EXFRAME:
   - Reads the ObjectInfo pointer to get project and module information
   - Reads the ObjectTable to extract project names
   - Extracts procedure/function names from the EXFRAME structure
6. Returns formatted callstack string without needing to change VBE execution state
7. **Works with both compiled (MDE) and uncompiled (MDB) VBA projects**
8. **No internet connection required** - operates entirely on in-memory structures

### Legacy Method (VBECallstackProvider - Obsolete)

1. On first use, VBAStack calls PdbEnum to extract function addresses from VBE7.DLL's PDB symbols
2. PdbEnum uses DbgHelp.dll to download symbols from Microsoft Symbol Servers if not already cached (~4MB), then searches for target functions in VBE7 
3. VBAStack caches the function pointers for subsequent calls
4. Calls "EbMode" from VBE7 to check the editor state
5. Sets up a window hook to prevent VBE window from flashing on screen
6. If VBE is in "Run" mode, switches it to "Break" mode using "EbSetMode" (required for callstack functions to work)
7. Calls "EbGetCallstackCount" to determine number of stack frames
8. Loops through each frame, calling "ErrGetCallstackString" to get formatted strings
9. Formats the results (converting "." to "::")
10. Restores original VBE mode and window state
11. Returns formatted callstack string

## Limitations

- **VBE 7.0+ Only**: Does not currently support Office 2007 or earlier (VBA6)
- **Windows Only**: Relies on Windows-specific APIs (DbgHelp.dll for legacy method, kernel32.dll)
- **Legacy Method Only:**
  - **PDB Dependency**: Requires Microsoft symbol servers to be accessible for first-time symbol download. This WILL NOT WORK in offline environments unless symbols are pre-cached.
  - **Symbol Version Matching**: If Microsoft updates VBE7.dll, symbols must match the DLL version or the legacy method will fail.
  - **Performance**: First call may be slow due to symbol loading/downloading (typically under 1 second locally, but depends on internet speed for ~4MB symbol file download); subsequent calls are near instant.
- **Direct Method Benefits:**
  - No internet connection required
  - No symbol download needed
  - Works with compiled (MDE) and uncompiled (MDB) projects
  - Faster initial execution (no PDB loading delay)

## Building

The solution uses Visual Studio's new solution format (.slnx) and targets .NET Framework 4.8.

### Prerequisites
- Visual Studio 2022 or later
- .NET Framework 4.8 SDK
- C# 9.0 support (for NativePtrCaller project)

### Build Steps
1. Open `VBAStack.slnx` in Visual Studio
2. Restore NuGet packages
3. Build the solution (Debug or Release configuration)

The solution includes:
- VBAStack (VB.NET) - Main library
- PdbEnumBase (C#) - Symbol enumeration base library
- PdbEnum_x64 (C#) - x64 symbol enumeration executable
- PdbEnum_x86 (C#) - x86 symbol enumeration executable
- NativePtrCaller (C#) - Native function pointer wrapper
- PdbEnum.Tests (C#) - Unit tests using MSTest framework

### Running Tests
Run the PdbEnum.Tests project using Visual Studio Test Explorer or:
```
dotnet test
```

## Deployment

### Using NuGet Package (Recommended)

Add the [NuGet package](https://www.nuget.org/packages/VBAStack) to your VSTO or COM add-in project:
```
Install-Package VBAStack
```

The recommended direct method (VBEDirectCallstackReader) has no additional deployment requirements.

### Manual Deployment

If distributing manually:

#### For Direct Method (Recommended):
1. Include `VBAStack.dll` in your project references
2. No additional dependencies required - works offline

#### For Legacy Method (Not Recommended):
1. Include `VBAStack.dll` in your project references
2. Ensure `PdbEnum_x64.exe` and `PdbEnum_x86.exe` are in the same directory or a subdirectory
3. Include `NativePtrCaller.dll` and `PdbEnum.dll`
4. Ensure the target machine has internet access for initial PDB symbol download (~4MB)

## Troubleshooting

### "VBE version is less than 7.0"
Upgrade to Office 2010 or later, or try your hand at implementing this for VBA6 and remove that check.

### Empty callstack returned
- VBA code is not currently executing
- The VBE is in Design mode with no code on the stack
- Try using the direct method (VBEDirectCallstackReader) if you were using the legacy method

### Legacy Method Specific Issues

#### "Could not get pointers to necessary VBE7 functions"
- Check that PdbEnum executables (PdbEnum_x64.exe, PdbEnum_x86.exe) are present
- Verify internet connectivity for symbol server access
- Try using the direct method instead (VBEDirectCallstackReader)

#### Symbol download errors
- Ensure firewall allows access to Microsoft Symbol Servers
- Check if symbols are blocked by antivirus software
- Consider switching to the direct method which doesn't require symbols

## Migration from Legacy to Direct Method

If you're currently using VBECallstackProvider, migrating to VBEDirectCallstackReader is simple:

**Before:**
```vb.net
Dim callstack = VBECallstackProvider.GetCallstack(vbe)
```

**After:**
```vb.net
Dim callstack = VBEDirectCallstackReader.GetCallstackString()
```

Benefits of migration:
- No internet connection required
- Faster initial execution
- Works with compiled VBA (MDE files)
- More reliable (no dependency on symbol matching)

## Contributing

Contributions are welcome! This project includes:
- Unit tests in PdbEnum.Tests
- Logging via VBAStackLogger for debugging
- Documentation in code comments

## License

See repository for license information.

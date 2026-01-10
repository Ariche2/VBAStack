# VBAStack

A library for retrieving VBA callstack information at runtime from Office applications. This enables debugging and error reporting capabilities for VBA add-ins and COM add-ins targeting Microsoft Office.

## Overview

VBAStack allows you to programmatically retrieve the current VBA call stack from a running VBA application (Excel, Word, Access, etc.). It does need to be loaded into the same process as the VBE, as it calls internal undocumented functions inside VBE7.dll. As such, it's best deployed as part of a COM add-in or VBA add-in and called from VBA via the add-in.
I am personally using it with a VSTO addin for MS Access, which makes deployment incredibly easy (not quite as easy as a certain other tool that can get the callstack, but I'm working on it).
These functions are not publicly documented or even exported from the DLL, however Microsoft does include them in the symbol files available on the Microsoft Symbol Servers. As such, we can dynamically download the right PDB file and retrieve the necessary function pointers from it using DbgHelp.dll (the same way debuggers retrieve symbols).

Due to the functions undocumented nature, and my lack of confidence in my own ability, this library is provided as-is without any guarantees. Use at your own risk.

## Requirements

- **VBE 7.0+** (Office 2010 and later)
- **.NET Framework 4.8**

## Architecture

The solution consists of several interconnected projects:

### Core Projects

#### VBAStack (VB.NET)
The main library that provides the high-level API for retrieving VBA callstacks.

**Key Components:**
- `VBECallstackProvider` - Public API for getting callstack information
- `VBESymbolResolver` - Communicates with PdbEnum to resolve function addresses
- `VBENativeWrapper` - Marshals calls to native VBE7.DLL functions
- `VBEWindowHook` - Manages VBE window visibility during callstack capture
- `VBEEnums` - Defines VBE execution modes and constants

#### PdbEnum (C#)
Base library for enumerating symbols from PDB (Program Database) files. I do not recommend using this outside of this project, as I haven't tested it at all outside of getting the necessary symbols from VBE7.dll.

**Key Components:**
- `SymbolEnumerator` - Interfaces with DbgHelp.dll to load and enumerate symbols
- `ModuleHelper` - Manages module loading and symbol resolution
- `OutputFormatter` - Formats symbol information for consumption
- `ImageHlpStructures` - P/Invoke structures for DbgHelp API

#### PdbEnum_x64 / PdbEnum_x86 (C#)
These are external platform-specific console executables that extract symbol addresses from VBE7.DLL's PDB files. This is somewhat of a hack - use of DbgHelp.dll can mess with debugging, so I figured it was safer to isolate it to a separate process.

#### NativePtrCaller (C#)
Provides wrappers for calling native function pointers from managed code using C# 9.0 function pointers.

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
    callstack = Application.COMAddins("MyAddin").Object.VSTO_GetVbaCallstack()
    'or...
    callstack = Application.COMAddins("MyAddin").Object.Generic_GetVbaCallstack(Application.VBE)
    MsgBox "An error occurred. Callstack:" & vbCrLf & callstack
    Exit Sub
End Sub
```

VB.Net - within a VSTO or COM add-in with access to the VBE object
```vb.net
Imports VBAStack

'Expose a public function for VBA to call...
Public Function VSTO_GetVbaCallstack() As String
    Dim vbe As Object = Globals.ThisAddIn.Application.VBE ' "Globals.ThisAddin" is a VSTO thing
    Return VBECallstackProvider.GetCallstack(vbe)
End Function

Public Function Generic_GetVbaCallstack(vbe As Object) As String
    Return VBECallstackProvider.GetCallstack(vbe)
End Function
```
C# - same as above, but in C#
```csharp
using VBAStack;

//Expose a public function for VBA to call...
public string VSTO_GetVbaCallstack()
{
    Object vbe = Globals.ThisAddIn.Application.VBE; // "Globals.ThisAddin" is a VSTO thing
    return VBECallstackProvider.GetCallstack(vbe);
}
public string Generic_GetVbaCallstack(Object vbe)
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

1. On first use, VBAStack calls PdbEnum to extract function addresses from VBE7.DLL's PDB symbols
2. PdbEnum uses DbgHelp.dll to download symbols from Microsoft Symbol Servers if not already cached, then searches for our target functions in VBE7 
3. VBAStack then caches the function pointers for subsequent calls
4. It then calls "EbMode", from VBE7 - this checks what state the editor is in.
5. To make sure the VBE window doesn't flash up on screen during the next step, it sets up a window hook to intercept messages to the VBE, and hide it if necessary
6. If the VBE is in "Run" mode, we switch it to "Break" mode using "EbSetMode" - this is necessary for the next 2 steps, as otherwise the callstack functions won't work
7. We call "EbGetCallstackCount" to find out how many stack frames there are
8. We loop through each stack frame, calling "ErrGetCallstackString" to get a formatted string for each frame (which differs from the old versions of these functions - for VBA6 it seems you used EbGetCallstackFunction, but that returns rubbish in VBE7 (some kind of ID, I think). ErrGetCallstackString does call it itself though)
9. Do a little formatting on the results, because I prefer "::" to "."
10. Restores original VBE mode and window state, if we had to change them
11. Returns formatted callstack string

## Limitations

- **VBE 7.0+ Only**: Does not support currently Office 2007 or earlier (VBA6)
- **Windows Only**: Relies on Windows-specific APIs (DbgHelp.dll, kernel32.dll)
- **PDB Dependency**: Requires Microsoft symbol servers to be accessible for first-time symbol download. This WILL NOT WORK in offline environments unless symbols are pre-cached - and at that point, if Microsoft pushes any changes, it will no longer work as the symbols won't match the DLL.
- **Performance**: First call may be a touch slow due to symbol loading (usually under a second on my machine, if it needs to download symbols though its entirely dependent on internet speed. The symbol file is roughly ~4mb); subsequent calls are near instant

## Building

Todo. This sucks at the minute.

## Deployment

When distributing:

Just add the Nuget package to your project. Or;

1. Include `VBAStack.dll` in your project references
2. Ensure `PdbEnum_x64.exe` and `PdbEnum_x86.exe` are in the same directory or a subdirectory
3. Include `NativePtrCaller.dll` and `PdbEnum.dll`
4. Ensure the target machine has internet access for initial PDB symbol download

## Troubleshooting

### "VBE version is less than 7.0"
Upgrade to Office 2010 or later, or try your hand at implementing this for VBA6 and remove that check.

### "Could not get pointers to necessary VBE7 functions"
- Check that PdbEnum executables are present
- Verify internet connectivity for symbol server access

### Empty callstack returned
This may occur if:
- VBA code is not currently executing
- The VBE is in Design mode with no code on the stack
- There's an issue with symbol resolution

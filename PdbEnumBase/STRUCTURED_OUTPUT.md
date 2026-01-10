# PdbEnum - Structured Output Documentation

## Overview

PdbEnum supports multiple output formats for easy parsing by other programs:
- **Human** (default) - Human-readable text format
- **JSON** - JavaScript Object Notation
- **XML** - Extensible Markup Language

## Command-Line Options

```
PdbEnum.exe [options] <ProcessID> <ModuleName> <SymbolName>

Options:
  -json       Output in JSON format
  -xml        Output in XML format
  -q, -quiet  Suppress informational messages (recommended with structured output)
```

## Usage Examples

### Human-readable output (default)
```bash
PdbEnum.exe 1234 ntdll.dll NtCreateFile
```

### JSON output
```bash
PdbEnum.exe -json -quiet 1234 ntdll.dll NtCreateFile
```

### XML output
```bash
PdbEnum.exe -xml -quiet 1234 ntdll.dll NtCreateFile
```

## Output Formats

### JSON Format

```json
{
  "Success": true,
  "Module": {
    "Name": "ntdll.dll",
    "FullPath": "C:\\Windows\\System32\\ntdll.dll",
    "BaseAddress": 140703249932288,
    "BaseAddressHex": "0x7FFE12340000",
    "Size": 2097152,
    "EntryPoint": 0,
    "EntryPointHex": "0x0"
  },
  "PdbInfo": {
    "PdbGuid": "12345678-1234-5678-90ab-cdef12345678",
    "PdbAge": 1,
    "PdbFileName": "C:\\Symbols\\ntdll.pdb\\...",
    "SymType": 3,
    "SymTypeName": "PDB"
  },
  "Symbol": {
    "Name": "NtCreateFile",
    "Address": 140703250001920,
    "AddressHex": "0x7FFE12350000",
    "Size": 64,
    "Flags": 0,
    "FlagsHex": "0x0",
    "Tag": 5
  }
}
```

When symbol is not found:
```json
{
  "Success": true,
  "Module": { ... },
  "PdbInfo": { ... },
  "Symbol": null
}
```

On error:
```json
{
  "Success": false,
  "ErrorMessage": "Module 'invalid.dll' not found in process 1234"
}
```

### XML Format

```xml
<?xml version="1.0" encoding="utf-8"?>
<SymbolSearchResult>
  <Module>
    <Name>ntdll.dll</Name>
    <FullPath>C:\Windows\System32\ntdll.dll</FullPath>
    <BaseAddress>140703249932288</BaseAddress>
    <Size>2097152</Size>
    <EntryPoint>0</EntryPoint>
  </Module>
  <PdbInfo>
    <PdbGuid>12345678-1234-5678-90ab-cdef12345678</PdbGuid>
    <PdbAge>1</PdbAge>
    <PdbFileName>C:\Symbols\ntdll.pdb\...</PdbFileName>
    <SymType>3</SymType>
  </PdbInfo>
  <Symbol>
    <Name>NtCreateFile</Name>
    <Address>140703250001920</Address>
    <Size>64</Size>
    <Flags>0</Flags>
    <Tag>5</Tag>
  </Symbol>
  <Success>true</Success>
</SymbolSearchResult>
```

## Quiet Mode

The `-quiet` or `-q` flag suppresses informational and debug messages, outputting only the structured result. This is recommended when piping output to another program:

```bash
# All debug messages go to stderr, structured output to stdout
PdbEnum.exe -json -quiet 1234 ntdll.dll NtCreateFile > output.json 2>debug.log
```

## Parsing Examples

### PowerShell - JSON
```powershell
$result = PdbEnum.exe -json -quiet 1234 ntdll.dll NtCreateFile | ConvertFrom-Json
if ($result.Success -and $result.Symbol) {
    Write-Host "Symbol found at: $($result.Symbol.AddressHex)"
}
```

### Python - JSON
```python
import subprocess
import json

result = subprocess.run(
    ['PdbEnum.exe', '-json', '-quiet', '1234', 'ntdll.dll', 'NtCreateFile'],
    capture_output=True, text=True
)
data = json.loads(result.stdout)
if data['Success'] and data['Symbol']:
    print(f"Symbol found at: {data['Symbol']['AddressHex']}")
```

### C# - XML
```csharp
using System.Xml.Serialization;

var serializer = new XmlSerializer(typeof(SymbolSearchResult));
using var reader = new StringReader(xmlOutput);
var result = (SymbolSearchResult)serializer.Deserialize(reader);
if (result.Success && result.Symbol != null)
{
    Console.WriteLine($"Symbol found at: 0x{result.Symbol.Address:X}");
}
```

## Exit Codes

- **0**: Success (symbol may or may not be found, check output)
- **1**: Error occurred (invalid arguments, process not found, etc.)

## Data Types Reference

### Numeric Values
- All addresses are 64-bit unsigned integers (ulong)
- Sizes are 32-bit unsigned integers (uint)
- Both decimal and hexadecimal representations provided in JSON

### Symbol Type Values
- 0 = None
- 1 = COFF
- 2 = CodeView
- 3 = PDB
- 4 = Export
- 5 = Deferred
- 6 = SYM
- 7 = DIA
- 8 = Virtual

**Note**: The `SymTypeName` field may show "Export" even when `SymType` is 3 (PDB) if the PDB file could not be loaded. This occurs when DbgHelp reports PDB type but `PdbGuid` is all zeros and `PdbFileName` is empty, indicating symbols were actually loaded from the export table.

### PDB Information

When `PdbGuid` is `00000000-0000-0000-0000-000000000000` (all zeros), it indicates:
- No PDB file was found or loaded
- Symbols are being read from the module's export table instead
- Limited symbol information is available (only exported functions)
- This is common for:
  - Third-party DLLs without public symbols
  - Older Microsoft components
  - 32-bit processes on 64-bit Windows (WOW64)

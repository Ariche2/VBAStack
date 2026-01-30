Imports System.IO
Imports System.Text
Imports System.Runtime.InteropServices
Imports System.Runtime.Serialization.Json
Imports PdbEnum

''' <summary>
''' Resolves VBE7.DLL function symbols using PdbEnum and manages symbol caching.
''' </summary>
Friend Class VBESymbolResolver
    Private Shared ReadOnly s_SymbolCache As New Dictionary(Of String, IntPtr)
    Private Shared s_VBE7ModuleBase As IntPtr
    Private Shared s_VBE7ModuleSize As Long
    Private Shared s_VBE7ModuleEnd As IntPtr
    Private Shared s_VBE7Path As String
    Private Shared s_PdbEnumPath As String
    Private Shared s_Initialized As Boolean
    Public Shared SymsToGet As New List(Of String) From {
        "EbMode",
        "EbSetMode",
        "EbGetCallstackCount",
        "ErrGetCallstackString"
    }

#Region "Win32 Imports"
    Private Declare Function GetModuleHandle Lib "kernel32.dll" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As IntPtr
    Private Declare Function GetModuleFileName Lib "kernel32.dll" Alias "GetModuleFileNameA" (ByVal hModule As IntPtr, ByVal lpFilename As StringBuilder, ByVal nSize As Integer) As Integer
    Private Declare Function GetModuleInformation Lib "psapi.dll" (hProcess As IntPtr, hModule As IntPtr, ByRef lpmodinfo As ModuleInfo, cb As UInteger) As Boolean
    <StructLayout(LayoutKind.Sequential)>
    Private Structure ModuleInfo
        Public lpBaseOfDll As IntPtr
        Public SizeOfImage As UInteger
        Public EntryPoint As IntPtr
    End Structure
#End Region

#Region "Initialization"
    ''' <summary>
    ''' Gets the base address of the VBE7.DLL module.
    ''' </summary>
    Public Shared Function GetVBE7ModuleBase() As IntPtr
        If s_VBE7ModuleBase = IntPtr.Zero Then
            s_VBE7ModuleBase = GetModuleHandle("VBE7.DLL")
            If s_VBE7ModuleBase = IntPtr.Zero Then
                Throw New Exception("Could not get VBE7.DLL module handle. Is the VBE loaded? Win32 Error: " & Marshal.GetLastWin32Error())
            End If

            Dim sb As New StringBuilder(260)
            If GetModuleFileName(s_VBE7ModuleBase, sb, sb.Capacity) <> 0 Then
                s_VBE7Path = sb.ToString()
            Else
                Throw New Exception("Could not get VBE7.DLL path. GetModuleFileName returned 0. Win32 Error: " & Marshal.GetLastWin32Error())
            End If

            Dim modInfo As ModuleInfo
            Dim currentProcess As IntPtr = Process.GetCurrentProcess().Handle
            If GetModuleInformation(currentProcess, s_VBE7ModuleBase, modInfo, CUInt(Marshal.SizeOf(GetType(ModuleInfo)))) Then
                s_VBE7ModuleSize = CLng(modInfo.SizeOfImage)
                s_VBE7ModuleEnd = IntPtr.Add(s_VBE7ModuleBase, CInt(s_VBE7ModuleSize))
            Else
                Throw New Exception("Could not get VBE7.DLL module information. Win32 Error: " & Marshal.GetLastWin32Error())
            End If

        End If

        Return s_VBE7ModuleBase
    End Function

    ''' <summary>
    ''' Locates the PdbEnum.exe executable.
    ''' </summary>
    Private Shared Function GetPdbEnumPath() As String
        If Not String.IsNullOrEmpty(s_PdbEnumPath) Then
            Return s_PdbEnumPath
        End If

        Dim assemblyPath As String = New Uri(Reflection.Assembly.GetExecutingAssembly().CodeBase).LocalPath
        Dim assemblyDir As String = Path.GetDirectoryName(assemblyPath)
        Dim ExecutableName As String
        If IntPtr.Size = 8 Then
            ExecutableName = "PdbEnum_x64.exe"
        Else
            ExecutableName = "PdbEnum_x86.exe"
        End If

        Dim foundFiles As String() = Directory.GetFiles(assemblyDir, ExecutableName.Replace(".exe", ".*"), SearchOption.AllDirectories)

        Dim returnpath As String = String.Empty
        Dim potentialreturnpath As String = String.Empty
        For Each filepath In foundFiles
            'Part of an ugly hack to stop ClickOnce/VSTO from freaking out about bitness mismatch when assemblies are compiled as AnyCPU but the two PdbEnum exe's are architecture-specific
            If Path.GetExtension(filepath).ToLowerInvariant = ".notexe" Then
                potentialreturnpath = filepath
            ElseIf Path.GetFileName(filepath).ToLowerInvariant = ExecutableName.ToLowerInvariant() Then
                returnpath = filepath
                s_PdbEnumPath = returnpath
                Exit For
            End If
        Next
        If String.IsNullOrEmpty(returnpath) AndAlso Not String.IsNullOrEmpty(potentialreturnpath) Then
            File.Copy(potentialreturnpath, potentialreturnpath.Replace(".notexe", ".exe"))
            s_PdbEnumPath = potentialreturnpath.Replace(".notexe", ".exe")
        End If

        Return s_PdbEnumPath
    End Function

    ''' <summary>
    ''' Initializes the symbol resolver by batch-loading all required symbols.
    ''' </summary>
    Public Shared Function Initialize() As Boolean
        If s_Initialized Then
            Return True
        End If

        GetVBE7ModuleBase()

        If String.IsNullOrEmpty(s_VBE7Path) OrElse Not File.Exists(s_VBE7Path) Then
            VBAStackLogger.LogError("[VBESymbolResolver] VBE7.DLL path not found")
            Return False
        End If

        Dim pdbEnumPath As String = GetPdbEnumPath()
        If String.IsNullOrEmpty(pdbEnumPath) Then
            VBAStackLogger.LogError("[VBESymbolResolver] PdbEnum.exe not found")
            Throw New Exception("PdbEnum.exe not found")
        End If

        For Each sym In SymsToGet
            If Not s_SymbolCache.ContainsKey(sym) OrElse s_SymbolCache(sym) = IntPtr.Zero Then
                s_Initialized = False
                Exit For
            End If
        Next

        If s_Initialized Then
            Return True
        End If

        ' Core symbols needed for basic operation
        Dim symbolNames As String() = SymsToGet.ToArray()
        Dim batchResult As BatchSymbolSearchResult = CallPdbEnumBatch(symbolNames)

        If Not batchResult.Success OrElse batchResult.Symbols Is Nothing Then
            VBAStackLogger.LogError($"[VBESymbolResolver] Batch symbol resolution failed: {batchResult.ErrorMessage}")
            Return False
        End If

        For Each symbolResult In batchResult.Symbols
            If symbolResult.Success AndAlso symbolResult.Symbol IsNot Nothing Then
                Dim symbolPtr As New IntPtr(CLng(symbolResult.Symbol.Address))
                s_SymbolCache(symbolResult.SearchedSymbolName) = symbolPtr
                VBAStackLogger.LogDebug($"[VBESymbolResolver] Resolved {symbolResult.SearchedSymbolName} -> 0x{symbolResult.Symbol.Address:X}")
            Else
                VBAStackLogger.LogWarning($"[VBESymbolResolver] Failed to resolve {symbolResult.SearchedSymbolName}")
            End If
        Next

        s_Initialized = True
        Return True
    End Function
#End Region

#Region "Symbol Resolution"
    ''' <summary>
    ''' Gets a pointer to a specific symbol, using cache when available.
    ''' </summary>
    Public Shared Function GetSymbolPointer(symbolName As String) As IntPtr
        If Not Initialize() Then
            Return IntPtr.Zero
        End If

        If s_SymbolCache.ContainsKey(symbolName) Then
            Return s_SymbolCache(symbolName)
        End If

        VBAStackLogger.LogDebug($"[VBESymbolResolver] Symbol {symbolName} not in cache, performing individual lookup")
        Dim result As SymbolSearchResult = CallPdbEnum(symbolName)

        If Not result.Success OrElse result.Symbol Is Nothing Then
            VBAStackLogger.LogError($"[VBESymbolResolver] Failed to resolve {symbolName}: {result.ErrorMessage}")
            Return IntPtr.Zero
        End If

        Dim symbolPtr As New IntPtr(CLng(result.Symbol.Address))
        s_SymbolCache(symbolName) = symbolPtr

        VBAStackLogger.LogDebug($"[VBESymbolResolver] Resolved {symbolName} -> 0x{result.Symbol.Address:X}")
        Return symbolPtr
    End Function

    ''' <summary>
    ''' Gets the symbol name at a specific address in VBE7.DLL.
    ''' </summary>
    Public Shared Function GetSymbolAtAddressBatch(addresses As List(Of IntPtr)) As List(Of Tuple(Of IntPtr, String))
        If Not Initialize() Then
            Return Nothing
        End If

        ' Verify the address is within the VBE7 module
        For Each address In addresses
            If Not IsAddressInModule(address) Then
                VBAStackLogger.LogDebug($"[VBESymbolResolver] Address 0x{address.ToInt64():X} is not in VBE7.DLL")
                Return Nothing
            End If
        Next

        Dim batchresult As BatchSymbolSearchResult = CallPdbEnumForAddressBatch(addresses)

        Dim retList As New List(Of Tuple(Of IntPtr, String))
        For Each result As SymbolSearchResult In batchresult.Symbols
            If Not result.Success OrElse result.Symbol Is Nothing Then
                VBAStackLogger.LogDebug($"[VBESymbolResolver] No symbol found at address 0x{result.Symbol.Address:X}")
            Else
                VBAStackLogger.LogDebug($"[VBESymbolResolver] Found symbol {result.Symbol.Name} at address 0x{result.Symbol.Address:X}")
                retList.Add(New Tuple(Of IntPtr, String)(New IntPtr(CLng(result.Symbol.Address)), result.Symbol.Name))
            End If
        Next

        Return retList
    End Function
#End Region

#Region "PdbEnum Communication"
    Private Shared Function CallPdbEnum(symbolName As String) As SymbolSearchResult
        Dim pdbEnumPath As String = GetPdbEnumPath()
        If String.IsNullOrEmpty(pdbEnumPath) Then
            Return New SymbolSearchResult With {
                .Success = False,
                .ErrorMessage = "PdbEnum.exe not found"
            }
        End If

        Dim currentProcess As Integer = Process.GetCurrentProcess().Id
        Dim arguments As String = $"-json -quiet {currentProcess} VBE7.DLL {symbolName}"

        Try
            VBAStackLogger.LogDebug($"[VBESymbolResolver] Calling PdbEnum with args: {arguments}")
            Dim psi As New ProcessStartInfo With {
                .FileName = pdbEnumPath,
                .arguments = arguments,
                .UseShellExecute = False,
                .RedirectStandardOutput = True,
                .RedirectStandardError = True,
                    .CreateNoWindow = True
                }

            Using proc As Process = Process.Start(psi)
                Dim output As String = proc.StandardOutput.ReadToEnd()
                Dim errorOutput As String = proc.StandardError.ReadToEnd()
                proc.WaitForExit()

                If Not String.IsNullOrEmpty(errorOutput) Then
                    VBAStackLogger.LogWarning($"[VBESymbolResolver stderr] {errorOutput}")
                End If

                If proc.ExitCode <> 0 Then
                    Return New SymbolSearchResult With {
                            .Success = False,
                            .ErrorMessage = $"PdbEnum exited with code {proc.ExitCode}"
                        }
                End If

                Return ParseJsonResult(output)
            End Using
        Catch ex As Exception
            VBAStackLogger.LogError($"[VBESymbolResolver] Error calling PdbEnum: {ex.Message}")
            Return New SymbolSearchResult With {
                    .Success = False,
                    .ErrorMessage = ex.Message
                }
        End Try
    End Function

    Private Shared Function CallPdbEnumBatch(symbolNames As String()) As BatchSymbolSearchResult
        Dim pdbEnumPath As String = GetPdbEnumPath()
        If String.IsNullOrEmpty(pdbEnumPath) Then
            Return New BatchSymbolSearchResult With {
                .Success = False,
                .ErrorMessage = "PdbEnum.exe not found"
            }
        End If

        Dim currentProcess As Integer = Process.GetCurrentProcess().Id
        Dim symbolArgs As String = String.Join(" ", symbolNames)
        Dim arguments As String = $"-json -quiet {currentProcess} VBE7.DLL {symbolArgs}"

        Try
            VBAStackLogger.LogDebug($"[VBESymbolResolver] Calling PdbEnum with args: {arguments}")
            Dim psi As New ProcessStartInfo With {
                .FileName = pdbEnumPath,
                .arguments = arguments,
                .UseShellExecute = False,
                .RedirectStandardOutput = True,
                .RedirectStandardError = True,
                .CreateNoWindow = True
            }

            Using proc As Process = Process.Start(psi)
                Dim output As String = proc.StandardOutput.ReadToEnd()
                Dim errorOutput As String = proc.StandardError.ReadToEnd()
                proc.WaitForExit()

                If Not String.IsNullOrEmpty(errorOutput) Then
                    VBAStackLogger.LogWarning($"[VBESymbolResolver stderr] {errorOutput}")
                End If

                If proc.ExitCode <> 0 Then
                    Return New BatchSymbolSearchResult With {
                        .Success = False,
                        .ErrorMessage = $"PdbEnum exited with code {proc.ExitCode}"
                    }
                End If

                Return ParseBatchJsonResult(output)
            End Using
        Catch ex As Exception
            VBAStackLogger.LogError($"[VBESymbolResolver] Error calling PdbEnum: {ex.Message}")
            Return New BatchSymbolSearchResult With {
                    .Success = False,
                    .ErrorMessage = ex.Message
                }
        End Try
    End Function

    Private Shared Function CallPdbEnumForAddressBatch(addresses As List(Of IntPtr)) As BatchSymbolSearchResult
        Dim pdbEnumPath As String = GetPdbEnumPath()
        If String.IsNullOrEmpty(pdbEnumPath) Then
            Return New BatchSymbolSearchResult With {
                .Success = False,
                .ErrorMessage = "PdbEnum.exe not found"
            }
        End If

        Dim currentProcess As Integer = Process.GetCurrentProcess().Id
        Dim addressHex As String = String.Join(" ", addresses.Select(Function(a) $"0x{a.ToInt64():X}"))
        Dim arguments As String = $"-json -quiet -addr {currentProcess} VBE7.DLL {addressHex}"

        Try
            VBAStackLogger.LogDebug($"[VBESymbolResolver] Calling PdbEnum with args: {arguments}")
            Dim psi As New ProcessStartInfo With {
                .FileName = pdbEnumPath,
                .arguments = arguments,
                .UseShellExecute = False,
                .RedirectStandardOutput = True,
                .RedirectStandardError = True,
                .CreateNoWindow = True
            }

            Using proc As Process = Process.Start(psi)
                Dim output As String = proc.StandardOutput.ReadToEnd()
                Dim errorOutput As String = proc.StandardError.ReadToEnd()
                proc.WaitForExit()

                If Not String.IsNullOrEmpty(errorOutput) Then
                    VBAStackLogger.LogWarning($"[VBESymbolResolver stderr] {errorOutput}")
                End If

                If proc.ExitCode <> 0 Then
                    Return New BatchSymbolSearchResult With {
                        .Success = False,
                        .ErrorMessage = $"PdbEnum exited with code {proc.ExitCode}"
                    }
                End If

                ' Parse the batch result and extract the first symbol
                Dim batchResult As BatchSymbolSearchResult = ParseBatchJsonResult(output)
                If batchResult.Success AndAlso batchResult.Symbols IsNot Nothing AndAlso batchResult.Symbols.Count > 0 Then
                    Return batchResult
                Else
                    Return New BatchSymbolSearchResult With {
                        .Success = False,
                        .ErrorMessage = "No symbol found at address"
                    }
                End If
            End Using
        Catch ex As Exception
            VBAStackLogger.LogError($"[VBESymbolResolver] Error calling PdbEnum: {ex.Message}")
            Return New BatchSymbolSearchResult With {
                .Success = False,
                .ErrorMessage = ex.Message
            }
        End Try
    End Function

    Private Shared Function ParseJsonResult(json As String) As SymbolSearchResult
        Try
            Using ms As New MemoryStream(Encoding.UTF8.GetBytes(json))
                Dim serializer As New DataContractJsonSerializer(GetType(SymbolSearchResult))
                Return CType(serializer.ReadObject(ms), SymbolSearchResult)
            End Using
        Catch ex As Exception
            VBAStackLogger.LogError($"[VBESymbolResolver] JSON parse error: {ex.Message}")
            Return New SymbolSearchResult With {
                    .Success = False,
                    .ErrorMessage = $"JSON parse error: {ex.Message}"
                }
        End Try
    End Function

    Private Shared Function ParseBatchJsonResult(json As String) As BatchSymbolSearchResult
        Try
            Using ms As New MemoryStream(Encoding.UTF8.GetBytes(json))
                Dim serializer As New DataContractJsonSerializer(GetType(BatchSymbolSearchResult))
                Return CType(serializer.ReadObject(ms), BatchSymbolSearchResult)
            End Using
        Catch ex As Exception
            VBAStackLogger.LogError($"[VBESymbolResolver] JSON parse error: {ex.Message}")
            Return New BatchSymbolSearchResult With {
                        .Success = False,
                        .ErrorMessage = $"JSON parse error: {ex.Message}"
                    }
        End Try
    End Function
#End Region

#Region "Pointer Verification"
    ''' <summary>
    ''' Verifies that a function pointer is valid by checking for: 0xCC (INT3 instruction) padding, 0x90 (NOP instruction) padding, or a preceding RET instruction.
    ''' </summary>
    Public Shared Function VerifyFunctionPointer(functionPtr As IntPtr, functionName As String) As Boolean

        Dim HasINT3PaddingBytes As Boolean = True
        Dim HasNOPPaddingBytes As Boolean = True
        Dim HasPrecedingRET As Boolean = True

        'Get the 5 bytes preceding the function pointer
        Dim precedingAddress As IntPtr = IntPtr.Subtract(functionPtr, 5)
        Dim buffer(4) As Byte

        'Check for 5x INT3 instruction (present in the most recent versions of VBE7)
        Marshal.Copy(precedingAddress, buffer, 0, 5)
        For i As Integer = 0 To 4
            If buffer(i) <> &HCC Then
                HasINT3PaddingBytes = False
                Exit For
            End If
        Next

        'Check for at least 2x NOP instruction (present before EbMode and EbGetCallstackCount in older versions of VBE7)
        If Not HasINT3PaddingBytes Then
            HasNOPPaddingBytes = True
            For i As Integer = 3 To 4
                If buffer(i) <> &H90 Then
                    HasNOPPaddingBytes = False
                    Exit For
                End If
            Next
        End If

        'Check for singular RET instruction (present before EbSetMode and ErrGetCallstackString in older versions of VBE7).
        'I am aware this is clutching at straws.
        If Not HasINT3PaddingBytes AndAlso Not HasNOPPaddingBytes Then
            HasPrecedingRET = buffer(4) = &HC3
        End If

        'Did any of the checks pass?
        If Not HasINT3PaddingBytes AndAlso Not HasNOPPaddingBytes AndAlso Not HasPrecedingRET Then
            VBAStackLogger.LogError($"[VBESymbolResolver] {functionName}: Function pointer verification failed at 0x{functionPtr.ToInt64():X}")
            Return False
        Else
            VBAStackLogger.LogDebug($"[VBESymbolResolver] {functionName}: Successfully verified 0xCC padding at 0x{functionPtr.ToInt64():X}")
            Return True
        End If

    End Function

    ''' <summary>
    ''' Verifies all critical function pointers are valid.
    ''' </summary>
    Public Shared Function VerifyAllPointers() As Boolean
        If Not Initialize() Then
            VBAStackLogger.LogError("[VBESymbolResolver] Failed to initialize symbol resolver")
            Return False
        End If

        Dim symbolNames As String() = SymsToGet.ToArray()
        Dim allValid As Boolean = True

        For Each symbolName In symbolNames
            If Not s_SymbolCache.ContainsKey(symbolName) Then
                VBAStackLogger.LogError($"[VBESymbolResolver] Symbol {symbolName} not found in cache")
                Return False
            End If

            Dim ptr As IntPtr = s_SymbolCache(symbolName)
            If ptr = IntPtr.Zero Then
                VBAStackLogger.LogError($"[VBESymbolResolver] Symbol {symbolName} has null pointer")
                Return False
            End If

            VBAStackLogger.LogDebug($"[VBESymbolResolver] {symbolName}: 0x{ptr.ToInt64():X}")
            allValid = allValid And VerifyFunctionPointer(ptr, symbolName)
        Next

        If allValid Then
            VBAStackLogger.LogInfo("[VBESymbolResolver] All function pointers successfully verified with 0xCC padding")
        Else
            VBAStackLogger.LogError("[VBESymbolResolver] Function pointer validation failed")
        End If

        Return allValid
    End Function

    Friend Shared Function IsAddressInModule(potentialAddress As IntPtr) As Boolean
        Return potentialAddress.ToInt64() >= GetVBE7ModuleBase.ToInt64() AndAlso potentialAddress.ToInt64() < s_VBE7ModuleEnd.ToInt64()
    End Function
#End Region
End Class

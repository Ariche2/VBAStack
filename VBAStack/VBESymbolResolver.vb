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
    Private Shared s_VBE7Path As String
    Private Shared s_PdbEnumPath As String
    Private Shared s_Initialized As Boolean

#Region "Win32 Imports"
    Private Declare Function GetModuleHandle Lib "kernel32.dll" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As IntPtr
    Private Declare Function GetModuleFileName Lib "kernel32.dll" Alias "GetModuleFileNameA" (ByVal hModule As IntPtr, ByVal lpFilename As StringBuilder, ByVal nSize As Integer) As Integer
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

        If s_SymbolCache.ContainsKey("EbMode") AndAlso
           s_SymbolCache.ContainsKey("EbSetMode") AndAlso
           s_SymbolCache.ContainsKey("EbGetCallstackCount") AndAlso
           s_SymbolCache.ContainsKey("ErrGetCallstackString") Then
            s_Initialized = True
            Return True
        End If

        Dim symbolNames As String() = {"EbMode", "EbSetMode", "EbGetCallstackCount", "ErrGetCallstackString"}
        Dim batchResult As BatchSymbolSearchResult = CallPdbEnumBatch(symbolNames)

        If Not batchResult.Success OrElse batchResult.Symbols Is Nothing Then
            VBAStackLogger.LogError($"[VBESymbolResolver] Batch symbol resolution failed: {batchResult.ErrorMessage}")
            Return False
        End If

        For Each symbolResult In batchResult.Symbols
            If symbolResult.Success AndAlso symbolResult.Symbol IsNot Nothing Then
                Dim symbolPtr As New IntPtr(CLng(symbolResult.Symbol.Address))
                s_SymbolCache(symbolResult.SearchedSymbolName) = symbolPtr
                VBAStackLogger.LogDebug($"[VBESymbolResolver] Resolved {symbolResult.SearchedSymbolName} -> {symbolResult.Symbol.Address:X}")
            Else
                VBAStackLogger.LogError($"[VBESymbolResolver] Failed to resolve {symbolResult.SearchedSymbolName}")
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

        VBAStackLogger.LogDebug($"[VBESymbolResolver] Resolved {symbolName} -> {result.Symbol.Address:X}")
        Return symbolPtr
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
                .Arguments = arguments,
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
                .Arguments = arguments,
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
            VBAStackLogger.LogError($"[VBESymbolResolver] {functionName}: Function pointer verification failed at 0x{functionPtr:X}")
            Return False
        Else
            VBAStackLogger.LogDebug($"[VBESymbolResolver] {functionName}: Successfully verified 0xCC padding at 0x{functionPtr:X}")
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

        Dim symbolNames As String() = {"EbMode", "EbSetMode", "EbGetCallstackCount", "ErrGetCallstackString"}
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

            VBAStackLogger.LogDebug($"[VBESymbolResolver] {symbolName}: 0x{ptr.ToString("X")}")
            allValid = allValid And VerifyFunctionPointer(ptr, symbolName)
        Next

        If allValid Then
            VBAStackLogger.LogInfo("[VBESymbolResolver] All function pointers successfully verified with 0xCC padding")
        Else
            VBAStackLogger.LogError("[VBESymbolResolver] Function pointer validation failed")
        End If

        Return allValid
    End Function
#End Region
End Class

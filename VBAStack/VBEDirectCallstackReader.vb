Imports System.Runtime.InteropServices
Imports System.Text

''' <summary>
''' Direct callstack reader that walks VBE internal structures.
''' Uses rtcErrObj to get the top of the EXFRAME linked list, then walks ObjectInfo/ObjectTable
''' structures to extract project, module, and function names. Works in both compiled (MDE) and uncompiled (MDB) VBA projects.
''' </summary>
Public Class VBEDirectCallstackReader
    Private Shared ExFrameTOS_GlobalVar As IntPtr
    Private Declare Function rtcErrObj Lib "VBE7" () As IntPtr

#Region "Initialization"
    ''' <summary>
    ''' Gets the pointer to the global g_ExFrameTOS variable by navigating from rtcErrObj.
    ''' </summary>
    Public Shared Function Initialize() As Boolean
        Try



            'Get memory location of top of the stack
            Dim errObj As IntPtr = rtcErrObj()

            'Offset 0x18 (x86) of VBAErr is a pointer to the global EbThread variable in VBE7
            Dim g_ebThread As IntPtr = Marshal.ReadIntPtr(errObj, IntPtr.Size * 6)

            'Offset 0x0C (x86) or 0x10 (x64) of EbThread is a pointer to the global ExFrameTOS variable in VBE7
            If IntPtr.Size = 4 Then
                ExFrameTOS_GlobalVar = g_ebThread + &HC
            Else
                ExFrameTOS_GlobalVar = g_ebThread + &H10
            End If

            Return True
        Catch ex As Exception
            VBAStackLogger.LogError($"[VBEDirectCallstackReader] Initialization failed: {ex.Message}")
            Return False
        End Try
    End Function
#End Region

#Region "Stack Walking"
    ''' <summary>
    ''' Gets the top-of-stack EXFRAME pointer
    ''' </summary>
    Public Shared Function GetExFrameTOS() As IntPtr
        If Not Initialize() Then
            Return IntPtr.Zero
        End If
        Try
            'Get the top of stack (current EXFRAME)
            Dim pExFrame As IntPtr = IntPtr.Zero
            Try
                pExFrame = Marshal.ReadIntPtr(ExFrameTOS_GlobalVar)
            Catch
            End Try
            Return pExFrame
        Catch ex As Exception
            VBAStackLogger.LogError($"[VBEDirectCallstackReader] Error getting ExFrameTOS: {ex.Message}")
            Return IntPtr.Zero
        End Try
    End Function

    ''' <summary>
    ''' Walks the callstack and returns detailed information for each frame
    ''' </summary>
    Public Shared Function GetCallstackFrames() As List(Of CallstackFrameInfo)
        Dim frames As New List(Of CallstackFrameInfo)

        Try
            'Get the top of stack
            Dim pExFrame As IntPtr = GetExFrameTOS()

            If pExFrame = IntPtr.Zero Then
                VBAStackLogger.LogDebug("[VBEDirectCallstackReader] Could not get top-of-stack ExFrame")
                Return frames
            End If
            VBAStackLogger.LogDebug($"[VBEDirectCallstackReader] Top-of-stack ExFrame is 0x{pExFrame.ToInt64():X}")

            ' Walk the EXFRAME linked list
            Dim index As Integer = 0

            Do
                Try
                    ' All EXFRAMEs represent regular function calls
                    Dim frameInfo As New CallstackFrameInfo

                    VBAStackLogger.LogDebug($"[VBEDirectCallstackReader] Frame {index}: pExFrame=0x{pExFrame.ToInt64():X}")

                    ' Extract names for this frame
                    If ExtractFrameNames(pExFrame, frameInfo) Then
                        frameInfo.IsValid = True
                    End If

                    frames.Add(frameInfo)

                    ' Move to next frame - read pNext field at offset 0
                    pExFrame = Marshal.ReadIntPtr(pExFrame)
                    index += 1
                Catch ex As Exception
                    VBAStackLogger.LogError($"[VBEDirectCallstackReader] Error reading frame {index}: {ex.Message}")
                    Exit Do
                End Try
            Loop Until pExFrame = IntPtr.Zero OrElse index >= 1000 ' Safety limit

            Return frames
        Catch ex As Exception
            VBAStackLogger.LogError($"[VBEDirectCallstackReader] Error walking callstack: {ex.Message}")
            Return frames
        End Try
    End Function
#End Region

#Region "Frame Name Extraction"
    ''' <summary>
    ''' Extracts project/module/function names from an ExFrame by reading various internal VBE structures.
    ''' Flow: EXFRAME -> RTMI -> ObjectInfo -> ObjectTable -> Public Object Descriptor -> Method Names
    ''' </summary>
    Private Shared Function ExtractFrameNames(pExFrame As IntPtr, frameInfo As CallstackFrameInfo) As Boolean
        Try
            Dim ExFrame As New ExFrame_AnyCPU(pExFrame)

            ' Step 1: Read RTMI pointer from EXFRAME (offset 0xC on x86, 0x18 on x64)
            'Dim pRtmi As IntPtr = Marshal.ReadIntPtr(pExFrame, If(IntPtr.Size = 8, &H18, &HC))
            Dim pRtmi As IntPtr = ExFrame.lpRTMI
            If pRtmi = IntPtr.Zero Then
                VBAStackLogger.LogDebug("[VBEDirectCallstackReader] RTMI is null")
                Return False
            End If
            VBAStackLogger.LogDebug($"[VBEDirectCallstackReader] RTMI at 0x{pRtmi.ToInt64():X}")

            Dim RTMI As New RTMI_AnyCPU(pRtmi)

            ' Step 2: Read ObjectInfo pointer from RTMI (offset 0x0)
            'Dim pObjectInfo As IntPtr = Marshal.ReadIntPtr(pRtmi, 0)
            Dim pObjectInfo As IntPtr = RTMI.lpObjectInfo
            If pObjectInfo = IntPtr.Zero Then
                VBAStackLogger.LogDebug("[VBEDirectCallstackReader] ObjectInfo is null")
                Return False
            End If
            VBAStackLogger.LogDebug($"[VBEDirectCallstackReader] ObjectInfo at 0x{pObjectInfo.ToInt64():X}")

            Dim ObjectInfo As New ObjectInfo_AnyCPU(pObjectInfo)

            ' Step 3: Read ObjectTable pointer from ObjectInfo (offset 0x4 on x86, 0x8 on x64)
            'Dim pObjectTable As IntPtr = Marshal.ReadIntPtr(pObjectInfo, If(IntPtr.Size = 8, &H8, &H4))
            Dim pObjectTable As IntPtr = ObjectInfo.lpObjectTable
            If pObjectTable = IntPtr.Zero Then
                VBAStackLogger.LogDebug("[VBEDirectCallstackReader] ObjectTable is null")
                Return False
            End If
            VBAStackLogger.LogDebug($"[VBEDirectCallstackReader] ObjectTable at 0x{pObjectTable.ToInt64():X}")

            Dim ObjectTable As New ObjectTable_AnyCPU(pObjectTable)

            ' Step 4: Read project name from ObjectTable (offset 0x40)
            'Dim pszProjectName As IntPtr = Marshal.ReadIntPtr(pObjectTable, &H40)
            Dim pszProjectName As IntPtr = ObjectTable.lpszProjectName
            If pszProjectName <> IntPtr.Zero Then
                frameInfo.ProjectName = Marshal.PtrToStringAnsi(pszProjectName)
                VBAStackLogger.LogDebug($"[VBEDirectCallstackReader] Project name: {frameInfo.ProjectName}")
            Else
                frameInfo.ProjectName = "[Unknown]"
            End If

            ' Step 5: Read Public Object Descriptor back-pointer from ObjectInfo (offset 0x18)
            'Dim pPublicObject As IntPtr = Marshal.ReadIntPtr(pObjectInfo, &H18)
            Dim pPublicObject As IntPtr = ObjectInfo.lpObject
            If pPublicObject = IntPtr.Zero Then
                VBAStackLogger.LogDebug("[VBEDirectCallstackReader] Public Object Descriptor is null")
                Return False
            End If
            VBAStackLogger.LogDebug($"[VBEDirectCallstackReader] Public Object Descriptor at 0x{pPublicObject.ToInt64():X}")

            Dim PublicObject As New PublicObjectDescriptor_AnyCPU(pPublicObject)

            ' Step 6: Read module name from Public Object Descriptor (offset 0x18)
            'Dim pszObjectName As IntPtr = Marshal.ReadIntPtr(pPublicObject, &H18)
            Dim pszObjectName As IntPtr = PublicObject.lpszObjectName
            If pszObjectName <> IntPtr.Zero Then
                frameInfo.ModuleName = Marshal.PtrToStringAnsi(pszObjectName)
                VBAStackLogger.LogDebug($"[VBEDirectCallstackReader] Module name: {frameInfo.ModuleName}")
            Else
                frameInfo.ModuleName = "[Unknown]"
            End If

            ' Step 7: Read lpMethods array pointer from ObjectInfo (offset 0x24)
            'Dim pMethods As IntPtr = Marshal.ReadIntPtr(pObjectInfo, &H24)
            Dim pMethods As IntPtr = ObjectInfo.lpMethods
            If pMethods = IntPtr.Zero Then
                VBAStackLogger.LogDebug("[VBEDirectCallstackReader] Methods array is null")
                frameInfo.FunctionName = "[Unknown]"
                Return True ' Still return true - we got project and module names
            End If

            ' Step 8: Read method count from Public Object Descriptor (offset 0x1C)
            'Dim methodCount As Integer = Marshal.ReadInt32(pPublicObject, &H1C)
            Dim methodCount As Integer = PublicObject.dwMethodCount
            VBAStackLogger.LogDebug($"[VBEDirectCallstackReader] Method count: {methodCount}")


            ' Step 9: Search through methods array to find matching RTMI pointer
            Dim methodIndex As Integer = -1
            For i As Integer = methodCount - 1 To 0 Step -1
                Dim pMethodRtmi As IntPtr = Marshal.ReadIntPtr(pMethods, i * IntPtr.Size)
                If pMethodRtmi = pRtmi Then
                    methodIndex = i
                    VBAStackLogger.LogDebug($"[VBEDirectCallstackReader] Found matching RTMI at index {i}")
                    Exit For
                End If
            Next

            If methodIndex = -1 Then
                VBAStackLogger.LogDebug("[VBEDirectCallstackReader] RTMI not found in methods array")
                frameInfo.FunctionName = "[Unknown]"
                Return True ' Still return true - we got project and module names
            End If

            ' Step 10: Read lpMethodNames pointer from Public Object Descriptor (offset 0x20)
            Dim pMethodNames As IntPtr = PublicObject.lpMethodNames
            If pMethodNames = IntPtr.Zero Then
                VBAStackLogger.LogDebug("[VBEDirectCallstackReader] Method names array is null")
                frameInfo.FunctionName = "[Unknown]"
                Return True ' Still return true - we got project and module names
            End If

            ' Step 11: Read function name from method names array using found index
            Dim pszMethodName As IntPtr = Marshal.ReadIntPtr(pMethodNames, methodIndex * IntPtr.Size)
            If pszMethodName <> IntPtr.Zero Then
                frameInfo.FunctionName = Marshal.PtrToStringAnsi(pszMethodName)
                VBAStackLogger.LogDebug($"[VBEDirectCallstackReader] Function name: {frameInfo.FunctionName}")
            Else
                frameInfo.FunctionName = "[Unknown]"
            End If

            Return True

        Catch ex As Exception
            VBAStackLogger.LogError($"[VBEDirectCallstackReader] Error in compiled extraction: {ex.Message}")
            VBAStackLogger.LogError(ex.StackTrace)
            Return False
        End Try
    End Function
#End Region

#Region "Public API"
    ''' <summary>
    ''' Gets a formatted callstack string by walking the internal structures directly
    ''' </summary>
    Public Shared Function GetCallstackString(Optional IncludeProject As Boolean = False) As String
        Dim frames As List(Of CallstackFrameInfo) = GetCallstackFrames()
        Dim result As New Text.StringBuilder()

        ' Reverse the list to show most recent call first
        frames.Reverse()

        For Each frame In frames
            If frame.IsValid Then
                result.AppendLine(frame.ToString(IncludeProject))
            End If
        Next

        Return result.ToString()
    End Function

    ''' <summary>
    ''' Gets the currently executing VBA function in the format "ModuleName::ProcedureName".
    ''' Returns the most recent frame from the callstack.
    ''' </summary>
    Public Shared Function GetCurrentFunction() As CallstackFrameInfo

        Try
            Dim frameInfo As New CallstackFrameInfo

            'Get the top of stack
            Dim pExFrame As IntPtr = GetExFrameTOS()

            If pExFrame = IntPtr.Zero Then
                VBAStackLogger.LogDebug("[VBEDirectCallstackReader] Could not get top-of-stack ExFrame")
                Return frameInfo
            End If
            VBAStackLogger.LogDebug($"[VBEDirectCallstackReader] Top-of-stack ExFrame is 0x{pExFrame.ToInt64():X}")

            ' Extract names for the current frame

            ExtractFrameNames(pExFrame, frameInfo)
            Return frameInfo

        Catch ex As Exception
            VBAStackLogger.LogError($"[VBEDirectCallstackReader] Error getting current function: {ex.Message}")
            Throw
        End Try
    End Function
#End Region

#Region "Debug Helpers"
    ''' <summary>
    ''' Dumps raw memory from a given address in hex dump format.
    ''' Useful for debugging structure layouts when translating x86 to x64.
    ''' </summary>
    ''' <param name="address">Starting address to dump</param>
    ''' <param name="length">Number of bytes to dump</param>
    ''' <returns>Formatted hex dump string</returns>
    Public Shared Function DumpMemory(address As IntPtr, length As Integer) As String
        Dim bytesPerLine As Integer = 16
        If address = IntPtr.Zero Then
            Return "[Null pointer]"
        End If

        If length <= 0 Then
            Return "[Invalid length]"
        End If

        Dim result As New StringBuilder()
        Dim currentAddr As Long = address.ToInt64()

        result.AppendLine($"Memory dump at 0x{currentAddr:X} ({length} bytes):")
        result.AppendLine()

        Try
            Dim allBytes(length - 1) As Byte
            Dim totalCounter As Integer = 0
            For offset As Integer = 0 To length - 1 Step bytesPerLine
                ' Address column
                result.Append($"0x{currentAddr + offset:X8}  ")

                ' Hex bytes column
                Dim bytesInLine As Integer = Math.Min(bytesPerLine, length - offset)
                Dim lineBytes(bytesInLine - 1) As Byte

                For i As Integer = 0 To bytesInLine - 1
                    Try
                        lineBytes(i) = Marshal.ReadByte(New IntPtr(currentAddr + offset + i))
                        result.Append($"{lineBytes(i):X2} ")
                    Catch ex As Exception
                        result.Append("?? ")
                        lineBytes(i) = 0
                    End Try
                    allBytes(totalCounter) = lineBytes(i)
                    totalCounter += 1

                    ' Add extra space based on pointer size for readability
                    If bytesPerLine = 16 Then
                        If IntPtr.Size = 4 Then
                            ' 32-bit: 4 sections of 4 bytes (space after indices 3, 7, 11)
                            If (i + 1) Mod 4 = 0 AndAlso i < 15 Then
                                result.Append(" ")
                            End If
                        Else
                            ' 64-bit: 2 sections of 8 bytes (space after index 7)
                            If i = 7 Then
                                result.Append(" ")
                            End If
                        End If
                    End If
                Next

                ' Pad if incomplete line
                If bytesInLine < bytesPerLine Then
                    Dim padding As Integer = (bytesPerLine - bytesInLine) * 3
                    If bytesPerLine = 16 Then
                        If IntPtr.Size = 4 Then
                            ' 32-bit: Account for remaining section separators (3 total separators at positions 4, 8, 12)
                            padding += 3 - (bytesInLine \ 4)
                        Else
                            ' 64-bit: Account for separator after 8th byte if not yet printed
                            If bytesInLine < 8 Then
                                padding += 1
                            End If
                        End If
                    End If
                    result.Append(New String(" "c, padding))
                End If

                ' ASCII column
                result.Append(" |")
                For i As Integer = 0 To bytesInLine - 1
                    Dim c As Char = ChrW(lineBytes(i))
                    If Char.IsControl(c) OrElse lineBytes(i) < 32 OrElse lineBytes(i) > 126 Then
                        result.Append("."c)
                    Else
                        result.Append(c)
                    End If
                Next
                result.AppendLine("|")
            Next

            Dim foundPointers As New List(Of Tuple(Of IntPtr, Integer))

            For i As Integer = 0 To length - IntPtr.Size
                Dim val As Long
                If IntPtr.Size = 8 Then
                    val = BitConverter.ToInt64(allBytes, i)
                Else
                    val = BitConverter.ToInt32(allBytes, i)
                End If
                Dim potentialPointer As New IntPtr(val)
                If VBESymbolResolver.IsAddressInModule(potentialPointer) Then
                    foundPointers.Add(New Tuple(Of IntPtr, Integer)(potentialPointer, i))
                End If
            Next

            If foundPointers.Count > 0 Then
                result.AppendLine()
                result.AppendLine($"VBE7 module base: 0x{VBESymbolResolver.GetVBE7ModuleBase.ToInt64():X}")
                result.AppendLine()
                result.AppendLine("Potential pointers found in dump:")

                Dim justPointers As New List(Of IntPtr)
                For Each ptr_offset_pair As Tuple(Of IntPtr, Integer) In foundPointers
                    justPointers.Add(ptr_offset_pair.Item1)
                Next

                Dim symbolInfoList As List(Of Tuple(Of IntPtr, String)) = VBESymbolResolver.GetSymbolAtAddressBatch(justPointers)

                For Each ptr_offset_pair As Tuple(Of IntPtr, Integer) In foundPointers
                    Dim symbolInfo = symbolInfoList.Find(Function(t) t.Item1 = ptr_offset_pair.Item1)
                    If symbolInfo IsNot Nothing Then
                        result.AppendLine($"Offset 0x{ptr_offset_pair.Item2:X}, 0x{ptr_offset_pair.Item1.ToInt64():X} -> {symbolInfo.Item2}")
                    Else
                        result.AppendLine($"Offset 0x{ptr_offset_pair.Item2:X}, 0x{ptr_offset_pair.Item1.ToInt64():X} -> [No symbol found]")
                    End If
                Next
            End If


        Catch ex As Exception
            result.AppendLine()
            result.AppendLine($"[Error reading memory: {ex.Message}]")
        End Try

        Return result.ToString()
    End Function
#End Region
End Class

''' <summary>
''' Represents a single frame in the callstack with parsed information
''' </summary>
Public Class CallstackFrameInfo
    Public Property ProjectName As String
    Public Property ModuleName As String
    Public Property FunctionName As String
    Public Property IsValid As Boolean

    Public Overrides Function ToString() As String
        Return $"{ProjectName}.{ModuleName}.{FunctionName}"
    End Function

    Public Overloads Function ToString(IncludeProject As Boolean) As String
        Dim str As String = ""
        If IncludeProject Then
            str &= ProjectName & "."
        End If
        str &= ModuleName & "::" & FunctionName
        Return str
    End Function
End Class

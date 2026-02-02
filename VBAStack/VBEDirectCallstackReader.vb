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


            ' Extra Step: Read function prototype from PrivateObject (offset 0x18 on x86, unknown on x64)
            Dim lpPrivateObject As IntPtr = ObjectInfo.lpPrivateObject
            If lpPrivateObject = IntPtr.Zero Then
                Return True
            End If
            'Only need one thing from PrivateObject so cba to setup a proper AnyCPU struct
            Dim pFuncProto As IntPtr = Marshal.ReadIntPtr(Marshal.ReadIntPtr(lpPrivateObject + IIf(IntPtr.Size = 4, &H18, &H999), methodIndex * IntPtr.Size)) 'TODO - need to find x64 offset - for now random value
            Dim FuncProto As New funcPrototype_AnyCPU(pFuncProto)
            VBAStackLogger.LogDebug($"[VBEDirectCallstackReader] Function Prototype at 0x{pFuncProto.ToInt64():X}, Param Count: {FuncProto.Params.Count}")

            ' Step 12: Extract parameters from the stack frame
            'TODO - get param sizes from FuncProto and properly parse param values
            'ExtractFrameParameters(pExFrame, frameInfo)


            Return True

        Catch ex As Exception
            VBAStackLogger.LogError($"[VBEDirectCallstackReader] Error in compiled extraction: {ex.Message}")
            VBAStackLogger.LogError(ex.StackTrace)
            Return False
        End Try
    End Function

    ''' <summary>
    ''' Extracts parameter values from the stack frame of an ExFrame
    ''' NOTE: This is EXPERIMENTAL and currently disabled due to accuracy issues.
    ''' The problem: argSz is stack cleanup size (bytes), not parameter count.
    ''' We need parameter metadata (count, types, ByVal/ByRef) to properly parse parameters.
    ''' This metadata exists in TYPE_DATA/EXTBL structures but requires deeper reverse engineering.
    ''' </summary>
    Private Shared Sub ExtractFrameParameters(pExFrame As IntPtr, frameInfo As CallstackFrameInfo)
        ' TODO: Implement proper parameter extraction by:
        ' 1. Finding parameter metadata in RTMI or related structures
        ' 2. Reading parameter count, types, and ByVal/ByRef flags
        ' 3. Parsing stack frame based on actual parameter layout
        ' For now, disable to avoid incorrect results
        Return

        Try
            Dim ExFrame As New ExFrame_AnyCPU(pExFrame)
            Dim pRtmi As IntPtr = ExFrame.lpRTMI
            If pRtmi = IntPtr.Zero Then
                Return
            End If

            Dim RTMI As New RTMI_AnyCPU(pRtmi)
            Dim argSz As UShort = RTMI.argSz
            Dim cbStackFrame As UShort = RTMI.cbStackFrame
            Dim cLocalVars As Integer = ExFrame.cLocalVars

            ' VBA VARIANTs are 16 bytes each
            Const VARIANT_SIZE As Integer = 16

            ' PROBLEM: This calculation is incorrect!
            ' argSz is stack cleanup size (total bytes), not parameter count
            ' Parameters can be different sizes depending on type and ByVal/ByRef
            Dim paramCount As Integer = 0
            If argSz > 0 Then
                ' This assumes all params are pointers, which is WRONG
                paramCount = argSz \ IntPtr.Size
            End If

            If paramCount = 0 Then
                Return
            End If

            VBAStackLogger.LogDebug($"[VBEDirectCallstackReader] EXPERIMENTAL: Extracting {paramCount} parameters (argSz={argSz}, cbStackFrame={cbStackFrame}, cLocalVars={cLocalVars})")

            ' Calculate the base address for local variables
            ' Local vars start at: (EXFRAME_address - 0x28 - cbStackFrame)
            Dim localVarBase As IntPtr = ExFrame.Address - &H28 - cbStackFrame

            ' Parameters are stored ABOVE the local variables in the stack frame
            ' They start after the local variables section
            Dim paramBase As IntPtr = localVarBase + (cLocalVars * VARIANT_SIZE)

            ' Read each parameter
            For i As Integer = 0 To paramCount - 1
                Try
                    ' Each parameter is a pointer to a VARIANT
                    Dim pVariant As IntPtr = Marshal.ReadIntPtr(paramBase + (i * IntPtr.Size))

                    If pVariant <> IntPtr.Zero Then
                        ' Read the VARIANT structure (first 16 bytes) safely
                        Dim variantData(VARIANT_SIZE - 1) As Byte

                        ' Use Marshal.Copy with error handling - it can throw AccessViolationException
                        ' if memory is invalid or protected
                        Try
                            Marshal.Copy(pVariant, variantData, 0, VARIANT_SIZE)
                        Catch ex As AccessViolationException
                            VBAStackLogger.LogError($"[VBEDirectCallstackReader] Cannot access memory for parameter {i} at 0x{pVariant.ToInt64():X}")
                            Continue For
                        End Try

                        ' First 2 bytes are the variant type (VT_*)
                        Dim vt As UShort = BitConverter.ToUInt16(variantData, 0)

                        ' Create parameter info
                        Dim param As New ParameterInfo With {
                            .Index = i,
                            .Address = pVariant,
                            .VariantType = vt,
                            .RawData = variantData
                        }

                        ' Try to extract the value based on variant type
                        param.Value = InterpretVariant(variantData, vt)

                        frameInfo.Parameters.Add(param)
                        VBAStackLogger.LogDebug($"[VBEDirectCallstackReader] Param {i}: VT={vt}, Value={param.Value}")
                    End If
                Catch ex As Exception
                    VBAStackLogger.LogError($"[VBEDirectCallstackReader] Error reading parameter {i}: {ex.Message}")
                End Try
            Next

        Catch ex As Exception
            VBAStackLogger.LogError($"[VBEDirectCallstackReader] Error extracting parameters: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' Interprets a VARIANT structure and returns a string representation of its value
    ''' </summary>
    Private Shared Function InterpretVariant(variantData() As Byte, vt As UShort) As String
        Try
            ' VARIANT layout: [0-1]=VT, [2-7]=reserved, [8-15]=data
            Select Case vt
                Case 0 ' VT_EMPTY
                    Return "<Empty>"
                Case 1 ' VT_NULL
                    Return "<Null>"
                Case 2 ' VT_I2 (short)
                    Return BitConverter.ToInt16(variantData, 8).ToString()
                Case 3 ' VT_I4 (long/int)
                    Return BitConverter.ToInt32(variantData, 8).ToString()
                Case 4 ' VT_R4 (float)
                    Return BitConverter.ToSingle(variantData, 8).ToString()
                Case 5 ' VT_R8 (double)
                    Return BitConverter.ToDouble(variantData, 8).ToString()
                Case 6 ' VT_CY (currency)
                    Dim cy As Long = BitConverter.ToInt64(variantData, 8)
                    Return (cy / 10000.0).ToString("C")
                Case 7 ' VT_DATE
                    Dim dateVal As Double = BitConverter.ToDouble(variantData, 8)
                    Return $"<Date: {dateVal}>"
                Case 8 ' VT_BSTR (string)
                    Dim pBstr As IntPtr = New IntPtr(BitConverter.ToInt32(variantData, 8))
                    If pBstr <> IntPtr.Zero Then
                        Try
                            Return $"""{Marshal.PtrToStringBSTR(pBstr)}"""
                        Catch
                            Return $"<BSTR at 0x{pBstr.ToInt64():X}>"
                        End Try
                    End If
                    Return "<Empty BSTR>"
                Case 9 ' VT_DISPATCH
                    Dim pDisp As IntPtr = New IntPtr(BitConverter.ToInt32(variantData, 8))
                    Return $"<IDispatch at 0x{pDisp.ToInt64():X}>"
                Case 10 ' VT_ERROR
                    Return $"<Error: 0x{BitConverter.ToInt32(variantData, 8):X}>"
                Case 11 ' VT_BOOL
                    Dim boolVal As Short = BitConverter.ToInt16(variantData, 8)
                    Return If(boolVal <> 0, "True", "False")
                Case 12 ' VT_VARIANT (should not occur at this level)
                    Return "<Variant>"
                Case 13 ' VT_UNKNOWN
                    Dim pUnk As IntPtr = New IntPtr(BitConverter.ToInt32(variantData, 8))
                    Return $"<IUnknown at 0x{pUnk.ToInt64():X}>"
                Case 16 ' VT_I1 (signed char)
                    Return CType(variantData(8), SByte).ToString()
                Case 17 ' VT_UI1 (byte)
                    Return variantData(8).ToString()
                Case 18 ' VT_UI2 (unsigned short)
                    Return BitConverter.ToUInt16(variantData, 8).ToString()
                Case 19 ' VT_UI4 (unsigned int)
                    Return BitConverter.ToUInt32(variantData, 8).ToString()
                Case 20 ' VT_I8 (long long)
                    Return BitConverter.ToInt64(variantData, 8).ToString()
                Case 21 ' VT_UI8 (unsigned long long)
                    Return BitConverter.ToUInt64(variantData, 8).ToString()
                Case &H2000 To &HFFFF ' VT_ARRAY flag
                    Return $"<Array: VT={vt:X}>"
                Case &H4000 To &HFFFF ' VT_BYREF flag
                    Dim pRef As IntPtr = New IntPtr(BitConverter.ToInt32(variantData, 8))
                    Return $"<ByRef VT={vt And &HFFF:X} at 0x{pRef.ToInt64():X}>"
                Case Else
                    Return $"<VT={vt} (0x{BitConverter.ToInt64(variantData, 8):X})>"
            End Select
        Catch ex As Exception
            Return $"<Error interpreting: {ex.Message}>"
        End Try
    End Function
#End Region

#Region "Public API"
    ''' <summary>
    ''' Gets a formatted callstack string by walking the internal structures directly
    ''' </summary>
    Public Shared Function GetCallstackString(Optional IncludeProject As Boolean = False, Optional includeParameters As Boolean = False) As String
        Dim frames As List(Of CallstackFrameInfo) = GetCallstackFrames()
        Dim result As New Text.StringBuilder()

        ' Reverse the list to show most recent call first
        frames.Reverse()

        For Each frame In frames
            If frame.IsValid Then
                If includeParameters AndAlso frame.Parameters.Count > 0 Then
                    result.AppendLine(frame.ToStringWithParameters())
                Else
                    result.AppendLine(frame.ToString(IncludeProject))
                End If
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

    ''' <summary>
    ''' Gets a detailed string representation of parameters for all frames in the callstack
    ''' </summary>
    Public Shared Function GetParametersDebugString() As String
        Dim result As New StringBuilder()
        
        Try
            Dim frames As List(Of CallstackFrameInfo) = GetCallstackFrames()
            
            If frames.Count = 0 Then
                Return "[No callstack frames found]"
            End If

            result.AppendLine($"VBA Callstack with Parameters ({frames.Count} frames):")
            result.AppendLine()

            For i As Integer = 0 To frames.Count - 1
                Dim frame As CallstackFrameInfo = frames(i)
                If frame.IsValid Then
                    result.AppendLine($"Frame {i}: {frame.ToString(True)}")
                    
                    If frame.Parameters.Count > 0 Then
                        For Each param In frame.Parameters
                            result.AppendLine($"    [{param.Index}] {param.Value} (VT={param.VariantType}, Addr=0x{param.Address.ToInt64():X})")
                        Next
                    Else
                        result.AppendLine("    (No parameters)")
                    End If
                    result.AppendLine()
                Else
                    result.AppendLine($"Frame {i}: [Invalid Frame]")
                    result.AppendLine()
                End If
            Next

        Catch ex As Exception
            result.AppendLine($"[Error: {ex.Message}]")
        End Try

        Return result.ToString()
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
    Public Property Parameters As New List(Of ParameterInfo)

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

    Public Function ToStringWithParameters() As String
        If Parameters.Count = 0 Then
            Return ToString()
        End If
        Dim paramStr As String = String.Join(", ", Parameters.Select(Function(p) p.Value))
        Return $"{ToString()}({paramStr})"
    End Function
End Class

''' <summary>
''' Represents a parameter extracted from a stack frame
''' </summary>
Public Class ParameterInfo
    Public Property Index As Integer
    Public Property Address As IntPtr
    Public Property VariantType As UShort
    Public Property RawData As Byte()
    Public Property Value As String

    Public Overrides Function ToString() As String
        Return $"[{Index}] {Value} (VT={VariantType})"
    End Function
End Class

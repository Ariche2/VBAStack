Imports System.Runtime.InteropServices

Module VBEStructures_AnyCPU
    Private Function AnyCPU_OffsetOf(struct As Type, Field As String) As Int32
        Dim returnVal As Int32 = Marshal.OffsetOf(struct, Field)
        If IntPtr.Size = 8 Then returnVal *= 2
        Return returnVal
    End Function
    Public Class ExFrame_AnyCPU
        Private pThis As IntPtr

        Public ReadOnly Property lpNext As IntPtr
            Get
                Try
                    Return Marshal.ReadIntPtr(pThis, AnyCPU_OffsetOf(GetType(EXFRAME), "lpNext"))
                Catch
                End Try
            End Get
        End Property

        Public ReadOnly Property lpRTMI As IntPtr
            Get
                Try
                    Return Marshal.ReadIntPtr(pThis, AnyCPU_OffsetOf(GetType(EXFRAME), "lpRTMI"))
                Catch
                End Try
            End Get
        End Property

        Public ReadOnly Property cLocalVars As Integer
            Get
                Try
                    Return Marshal.ReadInt32(pThis, AnyCPU_OffsetOf(GetType(EXFRAME), "cLocalVars"))
                Catch
                    Return 0
                End Try
            End Get
        End Property

        Public ReadOnly Property Address As IntPtr
            Get
                Return pThis
            End Get
        End Property

        Public Sub New(pExFrame As IntPtr)
            If pExFrame = IntPtr.Zero Then
                Throw New ArgumentNullException(NameOf(pExFrame))
            End If
            pThis = pExFrame
        End Sub
    End Class

    Public Class RTMI_AnyCPU
        Private pThis As IntPtr
        Public ReadOnly Property lpObjectInfo As IntPtr
            Get
                Try
                    ' lpObjectInfo is always at offset 0x00 on both platforms
                    Return Marshal.ReadIntPtr(pThis, 0)
                Catch
                    Return IntPtr.Zero
                End Try
            End Get
        End Property
        Public ReadOnly Property argSz As UShort
            Get
                Try
                    ' argSz is at offset 0x04 on x86, 0x08 on x64 (after the pointer)
                    Dim offset As Integer = If(IntPtr.Size = 8, &H8, &H4)
                    Return Marshal.ReadInt16(pThis, offset)
                Catch
                    Return 0
                End Try
            End Get
        End Property
        Public ReadOnly Property cbStackFrame As UShort
            Get
                Try
                    ' cbStackFrame is at offset 0x06 on x86, 0x0A on x64 (argSz + 2)
                    Dim offset As Integer = If(IntPtr.Size = 8, &HA, &H6)
                    Return Marshal.ReadInt16(pThis, offset)
                Catch
                    Return 0
                End Try
            End Get
        End Property
        Public Sub New(pRTMI As IntPtr)
            If pRTMI = IntPtr.Zero Then
                Throw New ArgumentNullException(NameOf(pRTMI))
            End If
            pThis = pRTMI
        End Sub
    End Class

    Public Class ObjectInfo_AnyCPU
        Private pThis As IntPtr
        Public ReadOnly Property lpObjectTable As IntPtr
            Get
                Try
                    Return Marshal.ReadIntPtr(pThis, AnyCPU_OffsetOf(GetType(OBJECTINFO), "lpObjectTable"))
                Catch
                    Return IntPtr.Zero
                End Try
            End Get
        End Property
        Public ReadOnly Property lpObject As IntPtr
            Get
                Try
                    Return Marshal.ReadIntPtr(pThis, AnyCPU_OffsetOf(GetType(OBJECTINFO), "lpObject"))
                Catch
                    Return IntPtr.Zero
                End Try
            End Get
        End Property
        Public ReadOnly Property lpPrivateObject As IntPtr
            Get
                Try
                    Return Marshal.ReadIntPtr(pThis, AnyCPU_OffsetOf(GetType(OBJECTINFO), "lpPrivateObject"))
                Catch
                    Return IntPtr.Zero
                End Try
            End Get
        End Property
        Public ReadOnly Property lpMethods As IntPtr
            Get
                Try
                    Return Marshal.ReadIntPtr(pThis, AnyCPU_OffsetOf(GetType(OBJECTINFO), "lpMethods"))
                Catch
                    Return IntPtr.Zero
                End Try
            End Get
        End Property
        Public Sub New(pObjectInfo As IntPtr)
            If pObjectInfo = IntPtr.Zero Then
                Throw New ArgumentNullException(NameOf(pObjectInfo))
            End If
            pThis = pObjectInfo
        End Sub
    End Class

    'This one doesn't just double the offsets for x64, so we have to special case it
    Public Class ObjectTable_AnyCPU
        Private pThis As IntPtr
        Public ReadOnly Property lpszProjectName As IntPtr
            Get
                Try
                    Return Marshal.ReadIntPtr(pThis, Marshal.OffsetOf(GetType(OBJECTTABLE), "lpszProjectName").ToInt32())
                Catch
                    Return IntPtr.Zero
                End Try
            End Get
        End Property
        Public Sub New(pObjectTable As IntPtr)
            If pObjectTable = IntPtr.Zero Then
                Throw New ArgumentNullException(NameOf(pObjectTable))
            End If
            pThis = pObjectTable
        End Sub
    End Class

    Public Class PublicObjectDescriptor_AnyCPU
        Private pThis As IntPtr
        Public ReadOnly Property lpszObjectName As IntPtr
            Get
                Try
                    Return Marshal.ReadIntPtr(pThis, AnyCPU_OffsetOf(GetType(PUBLIC_OBJECT_DESCRIPTOR), "lpszObjectName"))
                Catch
                    Return IntPtr.Zero
                End Try
            End Get
        End Property
        Public ReadOnly Property lpMethodNames As IntPtr
            Get
                Try
                    Return Marshal.ReadIntPtr(pThis, AnyCPU_OffsetOf(GetType(PUBLIC_OBJECT_DESCRIPTOR), "lpMethodNames"))
                Catch
                    Return IntPtr.Zero
                End Try
            End Get
        End Property
        Public ReadOnly Property dwMethodCount As IntPtr
            Get
                Try
                    Return Marshal.ReadIntPtr(pThis, AnyCPU_OffsetOf(GetType(PUBLIC_OBJECT_DESCRIPTOR), "dwMethodCount"))
                Catch
                    Return IntPtr.Zero
                End Try
            End Get
        End Property
        Public Sub New(pPublicObjectDescriptor As IntPtr)
            If pPublicObjectDescriptor = IntPtr.Zero Then
                Throw New ArgumentNullException(NameOf(pPublicObjectDescriptor))
            End If
            pThis = pPublicObjectDescriptor
        End Sub
    End Class

    Public Class funcPrototype_AnyCPU
        Public Structure ParamInfo
            Public Name As String
            Public baseTypeVal As eVBInternal_VarTypes
            Public isOptional As Boolean
            Public isArray As Boolean
            Public isByRef As Boolean
            Public hasExtraPointer As Boolean

            'Shamelessly copied from David Zimmer, https://www.gendigital.com/blog/insights/research/recovery-of-function-prototypes-in-visual-basic-6-executables
            Public Sub New(val As Byte, Name As String)
                '0x08 = 00001000  'long
                '0x28 = 00101000  'byref long
                '0x68 = 01101000  'array byref long    (0x40 -> 01000000)
                '0xA8 = 10101000  'optional byref long (0x80 -> 10000000)

                Me.Name = Name

                Dim tmp As Byte
                tmp = val

                If (tmp And &H80) = &H80 Then
                    Me.isOptional = True
                    tmp = tmp Xor &H80
                End If

                If (tmp And &H40) = &H40 Then
                    Me.isArray = True
                    tmp = tmp Xor &H40
                End If

                If (tmp And &H20) = &H20 Then
                    Me.isByRef = True
                    tmp = tmp Xor &H20
                End If

                Me.baseTypeVal = tmp

                If baseTypeVal = eVBInternal_VarTypes.epvT_ComIface Or baseTypeVal = eVBInternal_VarTypes.epvT_ComObj Or baseTypeVal = eVBInternal_VarTypes.epvT_Internal Then
                    Me.hasExtraPointer = True
                End If
            End Sub
        End Structure
        Private pThis As IntPtr
        Public Params As List(Of ParamInfo) = New List(Of ParamInfo)()
        Public ReadOnly Property argSize As Byte
            Get
                Try
                    Return Marshal.ReadByte(pThis, Marshal.OffsetOf(GetType(funcPrototype), "argSize").ToInt32())
                Catch
                    Return 0
                End Try
            End Get
        End Property
        Public ReadOnly Property lpArgNamesArray As IntPtr
            Get
                Try
                    Return Marshal.ReadIntPtr(pThis, Marshal.OffsetOf(GetType(funcPrototype), "lpArgNamesArray").ToInt32())
                Catch
                    Return IntPtr.Zero
                End Try
            End Get
        End Property
        Private Sub ReadParams()
            Dim startOfParamBytes As IntPtr = IntPtr.Add(pThis, Marshal.OffsetOf(GetType(funcPrototype), "END_unkn").ToInt32() + 1)

            Dim curParamByte As IntPtr = startOfParamBytes

            Do
                'The byte array ends either when we hit a 0 byte, or when the address we would read the next byte from is equal to lpArgNamesArray
                If curParamByte.ToInt64() >= lpArgNamesArray.ToInt64() Then
                    Exit Do
                End If

                Dim typeVal As Byte
                Try
                    typeVal = Marshal.ReadByte(curParamByte)
                Catch
                End Try

                If typeVal = 0 Then
                    Exit Do
                End If

                Dim name As String = String.Empty
                Try
                    name = Marshal.PtrToStringAnsi(Marshal.ReadIntPtr(lpArgNamesArray, Me.Params.Count * IntPtr.Size))
                Catch
                End Try

                Dim paramInfo As New ParamInfo(typeVal, name)
                Me.Params.Add(paramInfo)

                If paramInfo.hasExtraPointer Then
                    'Skip the extra bytes from the pointer
                    curParamByte = IntPtr.Add(curParamByte, IntPtr.Size)
                Else
                    curParamByte = IntPtr.Add(curParamByte, 1)
                End If
            Loop

        End Sub

        Public Sub New(pfuncPrototype As IntPtr)
            If pfuncPrototype = IntPtr.Zero Then
                Throw New ArgumentNullException(NameOf(pfuncPrototype))
            End If
            pThis = pfuncPrototype

            ReadParams()

        End Sub
    End Class
End Module

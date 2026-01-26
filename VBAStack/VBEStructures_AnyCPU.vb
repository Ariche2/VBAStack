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
                    Return Marshal.ReadIntPtr(pThis, AnyCPU_OffsetOf(GetType(RTMI), "lpObjectInfo"))
                Catch
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
                End Try
            End Get
        End Property
        Public ReadOnly Property lpObject As IntPtr
            Get
                Try
                    Return Marshal.ReadIntPtr(pThis, AnyCPU_OffsetOf(GetType(OBJECTINFO), "lpObject"))
                Catch
                End Try
            End Get
        End Property
        Public ReadOnly Property lpMethods As IntPtr
            Get
                Try
                    Return Marshal.ReadIntPtr(pThis, AnyCPU_OffsetOf(GetType(OBJECTINFO), "lpMethods"))
                Catch
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
                End Try
            End Get
        End Property
        Public ReadOnly Property lpMethodNames As IntPtr
            Get
                Try
                    Return Marshal.ReadIntPtr(pThis, AnyCPU_OffsetOf(GetType(PUBLIC_OBJECT_DESCRIPTOR), "lpMethodNames"))
                Catch
                End Try
            End Get
        End Property
        Public ReadOnly Property dwMethodCount As IntPtr
            Get
                Try
                    Return Marshal.ReadIntPtr(pThis, AnyCPU_OffsetOf(GetType(PUBLIC_OBJECT_DESCRIPTOR), "dwMethodCount"))
                Catch
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
End Module

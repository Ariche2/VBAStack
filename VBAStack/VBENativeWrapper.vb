''' <summary>
''' Provides managed wrappers around native VBE7.DLL functions.
''' </summary>
Friend Class VBENativeWrapper
    Private Shared s_EbModePtr As IntPtr
    Private Shared s_EbSetModePtr As IntPtr
    Private Shared s_EbGetCallstackCountPtr As IntPtr
    Private Shared s_ErrGetCallstackStringPtr As IntPtr

#Region "Function Pointer Properties"
    Private Shared ReadOnly Property EbModePtr As IntPtr
        Get
            If s_EbModePtr = IntPtr.Zero Then
                s_EbModePtr = VBESymbolResolver.GetSymbolPointer("EbMode")
                If s_EbModePtr = IntPtr.Zero Then
                    Throw New Exception("Failed to get pointer for EbMode")
                End If
            End If
            Return s_EbModePtr
        End Get
    End Property

    Private Shared ReadOnly Property EbSetModePtr As IntPtr
        Get
            If s_EbSetModePtr = IntPtr.Zero Then
                s_EbSetModePtr = VBESymbolResolver.GetSymbolPointer("EbSetMode")
                If s_EbSetModePtr = IntPtr.Zero Then
                    Throw New Exception("Failed to get pointer for EbSetMode")
                End If
            End If
            Return s_EbSetModePtr
        End Get
    End Property

    Private Shared ReadOnly Property EbGetCallstackCountPtr As IntPtr
        Get
            If s_EbGetCallstackCountPtr = IntPtr.Zero Then
                s_EbGetCallstackCountPtr = VBESymbolResolver.GetSymbolPointer("EbGetCallstackCount")
                If s_EbGetCallstackCountPtr = IntPtr.Zero Then
                    Throw New Exception("Failed to get pointer for EbGetCallstackCount")
                End If
            End If
            Return s_EbGetCallstackCountPtr
        End Get
    End Property

    Private Shared ReadOnly Property ErrGetCallstackStringPtr As IntPtr
        Get
            If s_ErrGetCallstackStringPtr = IntPtr.Zero Then
                s_ErrGetCallstackStringPtr = VBESymbolResolver.GetSymbolPointer("ErrGetCallstackString")
                If s_ErrGetCallstackStringPtr = IntPtr.Zero Then
                    Throw New Exception("Failed to get pointer for ErrGetCallstackString")
                End If
            End If
            Return s_ErrGetCallstackStringPtr
        End Get
    End Property
#End Region

#Region "Native Function Wrappers"
    ''' <summary>
    ''' Gets the current execution mode of the VBE.
    ''' </summary>
    Public Shared Function GetMode() As EbMode
        Return CType(NativePtrCaller.NativePtrCaller.EbMode(EbModePtr), EbMode)
    End Function

    ''' <summary>
    ''' Sets the execution mode of the VBE.
    ''' </summary>
    Public Shared Sub SetMode(mode As EbMode)
        NativePtrCaller.NativePtrCaller.EbSetMode(EbSetModePtr, CInt(mode))
    End Sub

    ''' <summary>
    ''' Gets the number of items in the current callstack.
    ''' </summary>
    Public Shared Function GetCallstackCount() As Integer
        Dim count As Integer = 0
        Dim retVal As Integer = NativePtrCaller.NativePtrCaller.EbGetCallstackCount(EbGetCallstackCountPtr, count)

        If retVal <> 0 Then
            Throw New Exception("Unknown native code error in GetCallstackCount - " & retVal)
        End If

        Return count
    End Function

    ''' <summary>
    ''' Gets the callstack entry at the specified index.
    ''' </summary>
    Public Shared Function GetCallstackString(index As Integer) As String
        Dim callstackString As String = ""
        Dim mysteryNumber As Integer = 0

        Dim retVal As Integer = NativePtrCaller.NativePtrCaller.ErrGetCallstackString(
            ErrGetCallstackStringPtr, index, callstackString, mysteryNumber)

        If retVal <> 0 Then
            Throw New Exception("Unknown native code error in GetCallstackString - " & retVal)
        End If

        Return callstackString
    End Function
#End Region
End Class

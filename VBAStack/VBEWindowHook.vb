Imports System.Windows.Forms
Imports System.Runtime.InteropServices

''' <summary>
''' Provides a window procedure hook for the Visual Basic Editor (VBE) window to control its visibility.
''' This class intercepts Windows messages sent to the VBE window and manipulates them to keep the window
''' hidden or shown according to the desired state.
''' </summary>
''' <remarks>
''' The VBEWindowHook uses NativeWindow to subclass the VBE window and intercept WM_WINDOWPOSCHANGING messages,
''' which allows it to modify the show/hide flags before the window position changes are applied.
''' This provides more reliable control over the VBE window visibility than simply calling ShowWindow.
''' </remarks>
Public Class VBEWindowHook
    Inherits NativeWindow

    'TODO: Potentially, could create a normal WinForms window and parent the VBE window to it, then hide that - when we're done, set the parent to 0?

    ''' <summary>
    ''' Win32 API function to show or hide a window.
    ''' </summary>
    ''' <param name="hWnd">Handle to the window.</param>
    ''' <param name="nCmdShow">Controls how the window is to be shown (e.g., SW_HIDE, SW_SHOW, SW_SHOWNA).</param>
    ''' <returns>True if the window was previously visible; otherwise, False.</returns>
    <DllImport("user32.dll", SetLastError:=True)>
    Private Shared Function ShowWindow(hWnd As IntPtr, nCmdShow As Integer) As <MarshalAs(UnmanagedType.Bool)> Boolean
    End Function

    ''' <summary>
    ''' When True, allows window messages to be processed normally without interception.
    ''' Used temporarily when we need to explicitly show or hide the window via ShowWindow API.
    ''' </summary>
    Private Property AllowMessages As Boolean

    ''' <summary>
    ''' Indicates whether the VBE window should remain hidden.
    ''' When True, WM_WINDOWPOSCHANGING messages are modified to force the window to stay hidden.
    ''' </summary>
    Public Property ShouldBeHidden As Boolean

    ''' <summary>
    ''' Windows message constants for various window-related notifications.
    ''' </summary>
    Private Enum WindowMessages As Integer
        ''' <summary>Sent when a window is being activated or deactivated.</summary>
        WM_ACTIVATE = &H6
        ''' <summary>Sent after a window has gained the keyboard focus.</summary>
        WM_SETFOCUS = &H7
        ''' <summary>Sent when a window is about to be shown or hidden.</summary>
        WM_SHOWWINDOW = &H18
        ''' <summary>Sent when the size, position, or Z-order is about to change. This is the key message we intercept.</summary>
        WM_WINDOWPOSCHANGING = &H46
        ''' <summary>Sent when the size, position, or Z-order has changed.</summary>
        WM_WINDOWPOSCHANGED = &H47
        ''' <summary>Sent when the non-client area needs to be painted.</summary>
        WM_NCPAINT = &H85
        ''' <summary>Sent when the non-client area needs to be activated or deactivated.</summary>
        WM_NCACTIVATE = &H86
        ''' <summary>Sent when the application is about to be activated or deactivated.</summary>
        WM_ACTIVATEAPP = &H1C
        ''' <summary>Performs no operation (null message).</summary>
        WM_NULL = 0
        ''' <summary>Sent when window styles are about to change.</summary>
        WM_STYLECHANGING = &H7C
    End Enum

    ''' <summary>
    ''' Initializes a new instance of the VBEWindowHook class and attaches it to the specified VBE window.
    ''' </summary>
    ''' <param name="VbeHwnd">The handle (HWND) of the VBE window to hook.</param>
    Public Sub New(VbeHwnd As IntPtr)
        AssignHandle(VbeHwnd)
    End Sub

    ''' <summary>
    ''' Overrides the window procedure to intercept and modify window messages.
    ''' Specifically intercepts WM_WINDOWPOSCHANGING to control window visibility.
    ''' </summary>
    ''' <param name="m">The Windows message to process.</param>
    ''' <remarks>
    ''' When AllowMessages is False, this method intercepts WM_WINDOWPOSCHANGING messages and modifies
    ''' the WINDOWPOS structure to add or remove the SWP_HIDEWINDOW and SWP_SHOWWINDOW flags based on
    ''' the ShouldBeHidden property. This ensures the window maintains the desired visibility state
    ''' even when other code or Windows itself attempts to show or hide it.
    ''' </remarks>
    Protected Overrides Sub WndProc(ByRef m As Message)
        If Not AllowMessages Then
            Select Case m.Msg
                Case WindowMessages.WM_WINDOWPOSCHANGING
                    ' Unmarshal the WINDOWPOS structure from the message parameter
                    Dim pos As WINDOWPOS = Marshal.PtrToStructure(Of WINDOWPOS)(m.LParam)

                    If ShouldBeHidden Then
                        ' Force the window to be hidden by removing show flag and adding hide flag
                        pos.flags = pos.flags And Not WindowPosFlags.SWP_SHOWWINDOW
                        pos.flags = pos.flags Or WindowPosFlags.SWP_HIDEWINDOW
                    Else
                        ' Force the window to be shown by removing hide flag and adding show flag
                        pos.flags = pos.flags And Not WindowPosFlags.SWP_HIDEWINDOW
                        pos.flags = pos.flags Or WindowPosFlags.SWP_SHOWWINDOW
                    End If

                    ' Marshal the modified structure back to the message parameter
                    Marshal.StructureToPtr(pos, m.LParam, True)
            End Select
        End If

        ' Pass the message to the base window procedure
        MyBase.WndProc(m)
    End Sub

    ''' <summary>
    ''' Ensures the VBE window is hidden by calling the ShowWindow API directly.
    ''' </summary>
    ''' <remarks>
    ''' Temporarily sets AllowMessages to True to bypass message interception,
    ''' then calls ShowWindow with SW_HIDE (0) to hide the window.
    ''' This is used for initial hiding or when the window needs to be forcibly hidden.
    ''' </remarks>
    Public Sub EnsureHidden()
        Try
            ' Temporarily allow messages to pass through without interception
            AllowMessages = True
            ' SW_HIDE = 0: Hides the window and activates another window
            Dim SW_HIDE As Integer = 0
            ShowWindow(Me.Handle, SW_HIDE)
        Finally
            ' Re-enable message interception
            AllowMessages = False
        End Try
    End Sub

    ''' <summary>
    ''' Ensures the VBE window is shown by calling the ShowWindow API directly.
    ''' </summary>
    ''' <remarks>
    ''' Temporarily sets AllowMessages to True to bypass message interception,
    ''' then calls ShowWindow with SW_SHOWNA (8) to show the window without activating it.
    ''' This is used when the window needs to be made visible again.
    ''' </remarks>
    Public Sub EnsureShown()
        Try
            ' Temporarily allow messages to pass through without interception
            AllowMessages = True
            ' SW_SHOWNA = 8: Shows the window in its current size and position without activating it
            Dim SW_SHOWNA As Integer = 8
            ShowWindow(Me.Handle, SW_SHOWNA)
        Finally
            ' Re-enable message interception
            AllowMessages = False
        End Try
    End Sub

    ''' <summary>
    ''' Contains information about the size and position of a window.
    ''' This structure is used with the WM_WINDOWPOSCHANGING and WM_WINDOWPOSCHANGED messages.
    ''' </summary>
    ''' <remarks>
    ''' The StructLayout attribute ensures the fields are laid out in memory exactly as defined,
    ''' matching the Win32 WINDOWPOS structure.
    ''' </remarks>
    <Runtime.InteropServices.StructLayout(Runtime.InteropServices.LayoutKind.Sequential)>
    Private Structure WINDOWPOS
        ''' <summary>Handle to the window.</summary>
        Public hwnd As IntPtr
        ''' <summary>Handle to the window behind which this window is placed (Z-order).</summary>
        Public hwndInsertAfter As IntPtr
        ''' <summary>Specifies the position of the left edge of the window.</summary>
        Public x As Integer
        ''' <summary>Specifies the position of the top edge of the window.</summary>
        Public y As Integer
        ''' <summary>Specifies the window width, in pixels.</summary>
        Public cx As Integer
        ''' <summary>Specifies the window height, in pixels.</summary>
        Public cy As Integer
        ''' <summary>Specifies the window position flags.</summary>
        Public flags As WindowPosFlags
    End Structure

    ''' <summary>
    ''' Flags used with the WINDOWPOS structure to specify window positioning options.
    ''' These flags can be combined using bitwise OR operations.
    ''' </summary>
    <Flags>
    Private Enum WindowPosFlags As UInteger
        ''' <summary>Retains the current size (ignores the cx and cy parameters).</summary>
        SWP_NOSIZE = 1
        ''' <summary>Retains the current position (ignores the x and y parameters).</summary>
        SWP_NOMOVE = 1 << 1
        ''' <summary>Retains the current Z order (ignores the hwndInsertAfter parameter).</summary>
        SWP_NOZORDER = 1 << 2
        ''' <summary>Does not redraw changes.</summary>
        SWP_NOREDRAW = 1 << 3
        ''' <summary>Does not activate the window.</summary>
        SWP_NOACTIVATE = 1 << 4
        ''' <summary>Sends a WM_NCCALCSIZE message to the window, even if the window's size is not changing.</summary>
        SWP_FRAMECHANGED = 1 << 5
        ''' <summary>Displays the window.</summary>
        SWP_SHOWWINDOW = 1 << 6
        ''' <summary>Hides the window.</summary>
        SWP_HIDEWINDOW = 1 << 7
        ''' <summary>Discards the entire contents of the client area.</summary>
        SWP_NOCOPYBITS = 1 << 8
        ''' <summary>Does not change the owner window's position in the Z order.</summary>
        SWP_NOREPOSITION = 1 << 9
        ''' <summary>Prevents the window from receiving the WM_WINDOWPOSCHANGING message.</summary>
        SWP_NOSENDCHANGING = 1 << 10
    End Enum
End Class

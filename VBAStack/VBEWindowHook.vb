Imports System.Windows.Forms
Imports System.Runtime.InteropServices

Public Class VBEWindowHook
    Inherits NativeWindow

    <DllImport("user32.dll", SetLastError:=True)>
    Private Shared Function ShowWindow(hWnd As IntPtr, nCmdShow As Integer) As <MarshalAs(UnmanagedType.Bool)> Boolean
    End Function

    Private Property AllowMessages As Boolean
    Public Property ShouldBeHidden As Boolean
    Private Enum WindowMessages As Integer
        WM_ACTIVATE = &H6
        WM_SETFOCUS = &H7
        WM_SHOWWINDOW = &H18
        WM_WINDOWPOSCHANGING = &H46
        WM_WINDOWPOSCHANGED = &H47
        WM_NCPAINT = &H85
        WM_NCACTIVATE = &H86
        WM_ACTIVATEAPP = &H1C
        WM_NULL = 0
        WM_STYLECHANGING = &H7C
    End Enum
    Public Sub New(VbeHwnd As IntPtr)
        AssignHandle(VbeHwnd)
    End Sub
    Protected Overrides Sub WndProc(ByRef m As Message)
        If Not AllowMessages Then
            Select Case m.Msg
                Case WindowMessages.WM_WINDOWPOSCHANGING
                    Dim pos As WINDOWPOS = Marshal.PtrToStructure(Of WINDOWPOS)(m.LParam)
                    If ShouldBeHidden Then
                        pos.flags = pos.flags And Not WindowPosFlags.SWP_SHOWWINDOW
                        pos.flags = pos.flags Or WindowPosFlags.SWP_HIDEWINDOW
                    Else
                        pos.flags = pos.flags And Not WindowPosFlags.SWP_HIDEWINDOW
                        pos.flags = pos.flags Or WindowPosFlags.SWP_SHOWWINDOW
                    End If
                    Marshal.StructureToPtr(pos, m.LParam, True)
            End Select
        End If

        MyBase.WndProc(m)
    End Sub
    Public Sub EnsureHidden()
        Try
            AllowMessages = True
            Dim SW_HIDE As Integer = 0
            ShowWindow(Me.Handle, SW_HIDE)
        Finally
            AllowMessages = False
        End Try
    End Sub
    Public Sub EnsureShown()
        Try
            AllowMessages = True
            Dim SW_SHOWNA As Integer = 8
            ShowWindow(Me.Handle, SW_SHOWNA)
        Finally
            AllowMessages = False
        End Try
    End Sub

    <Runtime.InteropServices.StructLayout(Runtime.InteropServices.LayoutKind.Sequential)>
    Private Structure WINDOWPOS
        Public hwnd As IntPtr
        Public hwndInsertAfter As IntPtr
        Public x As Integer
        Public y As Integer
        Public cx As Integer
        Public cy As Integer
        Public flags As WindowPosFlags
    End Structure
    Private Enum WindowPosFlags As UInteger
        SWP_NOSIZE = 1
        SWP_NOMOVE = 1 << 1
        SWP_NOZORDER = 1 << 2
        SWP_NOREDRAW = 1 << 3
        SWP_NOACTIVATE = 1 << 4
        SWP_FRAMECHANGED = 1 << 5
        SWP_SHOWWINDOW = 1 << 6
        SWP_HIDEWINDOW = 1 << 7
        SWP_NOCOPYBITS = 1 << 8
        SWP_NOREPOSITION = 1 << 9
        SWP_NOSENDCHANGING = 1 << 10
    End Enum
End Class

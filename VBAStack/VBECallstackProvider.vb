''' <summary>
''' High-level API for retrieving VBA callstack information.
''' </summary>
Public Class VBECallstackProvider
    Private Shared s_IsVerified As Boolean

    ''' <summary>
    ''' Verifies that all required function pointers can be resolved and are valid.
    ''' </summary>
    Public Shared Function VerifyPointers() As Boolean
        If s_IsVerified Then
            Return True
        End If

        Try
            s_IsVerified = VBESymbolResolver.VerifyAllPointers()
        Catch ex As Exception
            Debug.WriteLine($"[VBECallstackProvider] Error verifying pointers: {ex.Message}")
            Debug.WriteLine(ex.StackTrace)
            s_IsVerified = False
        End Try

        Return s_IsVerified
    End Function

    ''' <summary>
    ''' Gets the current VBA callstack from the specified VBE object.
    ''' </summary>
    ''' <param name="vbe">A VBE object from an Office application.</param>
    ''' <returns>A formatted string containing the callstack, or an error message.</returns>
    Public Shared Function GetCallstack(vbe As Object) As String
        Dim originalVisibility As Boolean

        Try
            If CStr(vbe.Version).FirstOrDefault() <> "7"c Then
                Return "VBE version is less than 7.0 - VBAStack only works with VBE7+"
            Else
                originalVisibility = CBool(vbe.MainWindow.Visible)
            End If
        Catch ex As Exception
            Throw New Exception("Could not access VBE object. See inner exception.", ex)
        End Try

        Dim vbeHwnd As IntPtr
        Try
            vbeHwnd = New IntPtr(CLng(vbe.MainWindow.HWnd))
        Catch ex As Exception
            Throw New Exception("Could not get VBE window handle. See inner exception.", ex)
        End Try

        Dim hook As VBEWindowHook
        Try
            hook = New VBEWindowHook(vbeHwnd) With {
                .ShouldBeHidden = originalVisibility
            }
        Catch ex As Exception
            Throw New Exception("Could not hook VBE window. See inner exception.", ex)
        End Try

        If Not VerifyPointers() Then
            Throw New Exception("Could not get pointers to necessary VBE7 functions")
        End If

        Dim originalMode As EbMode = VBENativeWrapper.GetMode()

        If originalMode <> EbMode.Break Then
            VBENativeWrapper.SetMode(EbMode.Break)
        End If

        If VBENativeWrapper.GetMode() <> EbMode.Break Then
            Throw New Exception("Could not set EbMode to 'Break'")
        End If

        Dim callStackCount As Integer = VBENativeWrapper.GetCallstackCount()
        Dim result As New Text.StringBuilder()

        If callStackCount > 0 Then
            For i As Integer = callStackCount - 1 To 0 Step -1
                Try
                    Dim callStackStr As String = VBENativeWrapper.GetCallstackString(i)
                    Dim parts() As String = Split(callStackStr, ".")

                    If parts.Length >= 3 Then
                        Dim moduleName As String = parts(1)
                        Dim functionName As String = parts(2)
                        result.AppendLine($"{moduleName}::{functionName}")
                    Else
                        If callStackStr = "[<Non-Basic Code>]" Then
                            result.AppendLine(callStackStr)
                        Else
                            result.AppendLine($"Unparseable stack entry: {callStackStr}")
                        End If
                    End If
                Catch ex As Exception
                    result.AppendLine($"Could not read stack at index {i} - exception: {ex.Message}")
                End Try
            Next
        End If

        VBENativeWrapper.SetMode(originalMode)

        If originalVisibility = True Then
            hook.EnsureShown()
        Else
            hook.EnsureHidden()
        End If

        hook.ReleaseHandle()

        Return result.ToString()
    End Function
End Class

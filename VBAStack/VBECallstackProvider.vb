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
            VBAStackLogger.LogError($"[VBECallstackProvider] Error verifying pointers: {ex.Message}")
            VBAStackLogger.LogError(ex.StackTrace)
            s_IsVerified = False
        End Try

        Return s_IsVerified
    End Function

    ''' <summary>
    ''' Gets the current VBA callstack from the specified VBE object.
    ''' </summary>
    ''' <param name="vbe">A VBE object from an Office application.</param>
    ''' <returns>A formatted string containing the callstack, or an error message.</returns>
    Public Shared Function GetCallstack(vbe As Object, Optional ExcludeNonBasicCodeFrames As Boolean = False) As String
        Dim originalVisibility As Boolean

        ' Validate VBE version and check if it's currently visible
        Try
            If CStr(vbe.Version).FirstOrDefault() <> "7"c Then
                Return "VBE version is less than 7.0 - VBAStack only works with VBE7+"
            Else
                originalVisibility = CBool(vbe.MainWindow.Visible)
            End If
        Catch ex As Exception
            Throw New Exception("Could not access VBE object. See inner exception.", ex)
        End Try

        ' Get VBE window handle
        Dim vbeHwnd As IntPtr
        Try
            vbeHwnd = New IntPtr(CLng(vbe.MainWindow.HWnd))
        Catch ex As Exception
            Throw New Exception("Could not get VBE window handle. See inner exception.", ex)
        End Try

        ' Hook VBE window to control visibility
        Dim hook As VBEWindowHook
        Try
            hook = New VBEWindowHook(vbeHwnd) With {
                .ShouldBeHidden = originalVisibility
            }
        Catch ex As Exception
            Throw New Exception("Could not hook VBE window. See inner exception.", ex)
        End Try


        ' Go get function pointers, do our best to verify them so we don't hard crash
        If Not VerifyPointers() Then
            Throw New Exception("Could not get pointers to necessary VBE7 functions")
        End If

        ' Save the original mode to restore later
        Dim originalMode As EbMode = VBENativeWrapper.GetMode()

        If originalMode <> EbMode.Break Then
            ' Switch to Break mode to read the callstack
            VBENativeWrapper.SetMode(EbMode.Break)

            ' Verify we successfully switched modes
            If VBENativeWrapper.GetMode() <> EbMode.Break Then
                Throw New Exception("Could not set EbMode to 'Break'")
            End If
        End If

        ' Get number of frames on the stack
        Dim callStackCount As Integer = VBENativeWrapper.GetCallstackCount()
        Dim result As New Text.StringBuilder()

        ' Read each stack frame, from most recent to oldest, and add to result
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
                        ' Not sure why but on 64bit Office, there's one of these between each normal stack frame? Have an option to ignore them.
                        If callStackStr = "[<Non-Basic Code>]" Then
                            If Not ExcludeNonBasicCodeFrames Then
                                result.AppendLine(callStackStr)
                            End If
                        Else
                            result.AppendLine($"Unparseable stack entry: {callStackStr}")
                        End If
                    End If
                Catch ex As Exception
                    result.AppendLine($"Exception when reading stackframe {i}: {ex.Message}")
                End Try
            Next
        End If

        ' Restore original VBE mode
        VBENativeWrapper.SetMode(originalMode)

        ' Restore original VBE window visibility
        If originalVisibility = True Then
            hook.EnsureShown()
        Else
            hook.EnsureHidden()
        End If

        ' Release the hook on the VBE window
        hook.ReleaseHandle()

        Return result.ToString()
    End Function
End Class

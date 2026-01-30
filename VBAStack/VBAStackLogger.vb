''' <summary>
''' Centralized logging module for VBAStack project.
''' </summary>
Public Module VBAStackLogger

    ''' <summary>
    ''' Log severity levels.
    ''' </summary>
    Public Enum LogLevel
        Debug
        Info
        Warning
        [Error]
    End Enum

    ''' <summary>
    ''' When True, only Error level messages are logged. When False, all messages are logged.
    ''' </summary>
    Public Property ErrorsOnly As Boolean = True

    ''' <summary>
    ''' Logs a debug message.
    ''' </summary>
    Public Sub LogDebug(message As String)
        Log(LogLevel.Debug, message)
    End Sub

    ''' <summary>
    ''' Logs an informational message.
    ''' </summary>
    Public Sub LogInfo(message As String)
        Log(LogLevel.Info, message)
    End Sub

    ''' <summary>
    ''' Logs a warning message.
    ''' </summary>
    Public Sub LogWarning(message As String)
        Log(LogLevel.Warning, message)
    End Sub

    ''' <summary>
    ''' Logs an error message.
    ''' </summary>
    Public Sub LogError(message As String)
        Log(LogLevel.Error, message)
    End Sub

    ''' <summary>
    ''' Core logging method that respects the ErrorsOnly setting.
    ''' </summary>
    Private Sub Log(level As LogLevel, message As String)
        If ErrorsOnly AndAlso level <> LogLevel.Error Then
            Return
        End If

        Dim prefix As String = String.Empty
        Select Case level
            Case LogLevel.Debug
                prefix = "[DEBUG]"
            Case LogLevel.Info
                prefix = "[INFO]"
            Case LogLevel.Warning
                prefix = "[WARNING]"
            Case LogLevel.Error
                prefix = "[ERROR]"
        End Select

        Debug.WriteLine($"{prefix} {message}")
    End Sub

End Module

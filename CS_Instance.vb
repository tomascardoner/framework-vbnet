Module CS_Instance

    Friend Function IsRunningUnderIDE() As Boolean
        Return System.Diagnostics.Debugger.IsAttached()
    End Function

End Module
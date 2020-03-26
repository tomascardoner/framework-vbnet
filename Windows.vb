Imports System.Runtime.InteropServices

Namespace CardonerSistemas
    Friend Module Windows

#Region "Misc Windows API Declarations"

        <DllImport("user32.dll")>
        Private Function GetWindowTextLength(ByVal hwnd As IntPtr) As Integer
        End Function

        <DllImport("user32.dll", SetLastError:=True)>
        Private Function GetWindowThreadProcessId(ByVal hwnd As IntPtr, ByRef lpdwProcessId As IntPtr) As Integer
        End Function

        <DllImport("user32.dll", CharSet:=CharSet.Auto, SetLastError:=True, ExactSpelling:=True)>
        Public Function IsWindow(ByVal hWnd As IntPtr) As <MarshalAs(UnmanagedType.Bool)> Boolean
        End Function

        ''' <summary>The GetForegroundWindow function returns a handle to the foreground window.</summary>
        ''' <returns>The return value is a handle to the foreground window. The foreground window can be NULL in certain circumstances, such as when a window is losing activation. </returns>
        <DllImport("user32.dll", SetLastError:=True)>
        Private Function GetForegroundWindow() As IntPtr
        End Function

        <DllImport("user32.dll")>
        Private Function SetForegroundWindow(ByVal hWnd As IntPtr) As Boolean
        End Function

#End Region

#Region "GetWindow"

        <DllImport("user32.dll", SetLastError:=True, CharSet:=CharSet.Auto)>
        Private Function GetWindow(ByVal hWnd As IntPtr, ByVal uCmd As UInt32) As IntPtr
        End Function

        Private Enum GetWindowType As UInteger
            GW_HWNDFIRST = 0
            GW_HWNDLAST = 1
            GW_HWNDNEXT = 2
            GW_HWNDPREV = 3
            GW_OWNER = 4
            GW_CHILD = 5
            GW_ENABLEDPOPUP = 6
        End Enum

#End Region

#Region "GetWindowLong"

        <DllImport("user32.dll", EntryPoint:="GetWindowLong")>
        Private Function GetWindowLongPtr32(ByVal hWnd As IntPtr, ByVal nIndex As Integer) As Integer
        End Function

        <DllImport("user32.dll", EntryPoint:="GetWindowLongPtr")>
        Private Function GetWindowLongPtr64(ByVal hWnd As IntPtr, ByVal nIndex As Integer) As Integer
        End Function

        ' This static method is required because Win32 does not support GetWindowLongPtr directly
        Private Function GetWindowLongPtr(ByVal hWnd As IntPtr, ByVal nIndex As Integer) As Integer
            If IntPtr.Size = 8 Then
                Return GetWindowLongPtr64(hWnd, nIndex)
            Else
                Return GetWindowLongPtr32(hWnd, nIndex)
            End If
        End Function

        Private Enum GWL
            GWL_WNDPROC = -4
            GWL_HINSTANCE = -6
            GWL_HWNDPARENT = -8
            GWL_STYLE = -16
            GWL_EXSTYLE = -20
            GWL_USERDATA = -21
            GWL_ID = -12
        End Enum

#End Region

#Region "GetWindowStyle"

        Private Function GetWindowStyle(ByVal hWnd As IntPtr) As Integer
            Return GetWindowLongPtr(hWnd, GWL.GWL_STYLE)
        End Function

        ''' <summary>
        ''' Window Styles.
        ''' The following styles can be specified wherever a window style is required. After the control has been created, these styles cannot be modified, except as noted.
        ''' </summary>
        <Flags()> Private Enum WindowStyles As UInteger
            ''' <summary>The window has a thin-line border.</summary>
            WS_BORDER = &H800000

            ''' <summary>The window has a title bar (includes the WS_BORDER style).</summary>
            WS_CAPTION = &HC00000

            ''' <summary>The window is a child window. A window with this style cannot have a menu bar. This style cannot be used with the WS_POPUP style.</summary>
            WS_CHILD = &H40000000

            ''' <summary>Excludes the area occupied by child windows when drawing occurs within the parent window. This style is used when creating the parent window.</summary>
            WS_CLIPCHILDREN = &H2000000

            ''' <summary>
            ''' Clips child windows relative to each other; that is, when a particular child window receives a WM_PAINT message, the WS_CLIPSIBLINGS style clips all other overlapping child windows out of the region of the child window to be updated.
            ''' If WS_CLIPSIBLINGS is not specified and child windows overlap, it is possible, when drawing within the client area of a child window, to draw within the client area of a neighboring child window.
            ''' </summary>
            WS_CLIPSIBLINGS = &H4000000

            ''' <summary>The window is initially disabled. A disabled window cannot receive input from the user. To change this after a window has been created, use the EnableWindow function.</summary>
            WS_DISABLED = &H8000000

            ''' <summary>The window has a border of a style typically used with dialog boxes. A window with this style cannot have a title bar.</summary>
            WS_DLGFRAME = &H400000

            ''' <summary>
            ''' The window is the first control of a group of controls. The group consists of this first control and all controls defined after it, up to the next control with the WS_GROUP style.
            ''' The first control in each group usually has the WS_TABSTOP style so that the user can move from group to group. The user can subsequently change the keyboard focus from one control in the group to the next control in the group by using the direction keys.
            ''' You can turn this style on and off to change dialog box navigation. To change this style after a window has been created, use the SetWindowLong function.
            ''' </summary>
            WS_GROUP = &H20000

            ''' <summary>The window has a horizontal scroll bar.</summary>
            WS_HSCROLL = &H100000

            ''' <summary>The window is initially maximized.</summary>
            WS_MAXIMIZE = &H1000000

            ''' <summary>The window has a maximize button. Cannot be combined with the WS_EX_CONTEXTHELP style. The WS_SYSMENU style must also be specified.</summary>
            WS_MAXIMIZEBOX = &H10000

            ''' <summary>The window is initially minimized.</summary>
            WS_MINIMIZE = &H20000000

            ''' <summary>The window has a minimize button. Cannot be combined with the WS_EX_CONTEXTHELP style. The WS_SYSMENU style must also be specified.</summary>
            WS_MINIMIZEBOX = &H20000

            ''' <summary>The window is an overlapped window. An overlapped window has a title bar and a border.</summary>
            WS_OVERLAPPED = &H0

            ''' <summary>The window is an overlapped window.</summary>
            WS_OVERLAPPEDWINDOW = WS_OVERLAPPED Or WS_CAPTION Or WS_SYSMENU Or WS_SIZEFRAME Or WS_MINIMIZEBOX Or WS_MAXIMIZEBOX

            ''' <summary>The window is a pop-up window. This style cannot be used with the WS_CHILD style.</summary>
            WS_POPUP = &H80000000UI

            ''' <summary>The window is a pop-up window. The WS_CAPTION and WS_POPUPWINDOW styles must be combined to make the window menu visible.</summary>
            WS_POPUPWINDOW = WS_POPUP Or WS_BORDER Or WS_SYSMENU

            ''' <summary>The window has a sizing border.</summary>
            WS_SIZEFRAME = &H40000

            ''' <summary>The window has a window menu on its title bar. The WS_CAPTION style must also be specified.</summary>
            WS_SYSMENU = &H80000

            ''' <summary>
            ''' The window is a control that can receive the keyboard focus when the user presses the TAB key.
            ''' Pressing the TAB key changes the keyboard focus to the next control with the WS_TABSTOP style.  
            ''' You can turn this style on and off to change dialog box navigation. To change this style after a window has been created, use the SetWindowLong function.
            ''' For user-created windows and modeless dialogs to work with tab stops, alter the message loop to call the IsDialogMessage function.
            ''' </summary>
            WS_TABSTOP = &H10000

            ''' <summary>The window is initially visible. This style can be turned on and off by using the ShowWindow or SetWindowPos function.</summary>
            WS_VISIBLE = &H10000000

            ''' <summary>The window has a vertical scroll bar.</summary>
            WS_VSCROLL = &H200000
        End Enum

#End Region

#Region "GetWindowText"

        <DllImport("user32.dll", EntryPoint:="GetWindowText")>
        Public Function GetWindowText(ByVal hwnd As IntPtr, ByVal lpString As IntPtr, ByVal cch As Integer) As Integer
        End Function

        Private Function GetWindowText(ByVal hWnd As IntPtr, ByVal bTrim As Boolean) As String
            ' cc may be greater than the actual length;
            ' https://docs.microsoft.com/en-us/windows/desktop/api/winuser/nf-winuser-getwindowtextlengthw
            Dim cc As Integer = GetWindowTextLength(hWnd)

            If cc <= 0 Then
                Return String.Empty
            End If

            Dim strWindow As String
            Dim p As IntPtr = IntPtr.Zero
            Try
                Dim cbChar As Integer = Marshal.SystemDefaultCharSize
                Dim cb As Integer = (cc + 2) * cbChar
                p = Marshal.AllocCoTaskMem(cb)
                If p = IntPtr.Zero Then
                    Return String.Empty
                End If

                Dim pbZero(cb) As Byte
                Marshal.Copy(pbZero, 0, p, cb)

                Dim ccReal As Integer = GetWindowText(hWnd, p, cc + 1)
                If ccReal <= 0 Then
                    Return String.Empty
                End If

                If ccReal <= cc Then
                    ' Ensure correct termination (in case GetWindowText copied too much)
                    Dim ibZero As Integer = ccReal * cbChar
                    For i As Integer = 0 To cbChar - 1
                        Marshal.WriteByte(p, ibZero + i, 0)
                    Next i
                Else
                    Return String.Empty
                End If

                strWindow = Marshal.PtrToStringAuto(p)
                If strWindow = Nothing Then
                    strWindow = String.Empty
                End If

            Finally
                If p <> IntPtr.Zero Then
                    Marshal.FreeCoTaskMem(p)
                End If
            End Try

            Return CStr(IIf(bTrim, strWindow.Trim(), strWindow))
        End Function

#End Region

#Region "ChangeApplicationFocus"

        Friend Function LoseFocus(ByRef form As Form) As Boolean
            Try
                Dim sourceformhWnd As IntPtr = form.Handle
                Dim destinationhWnd As IntPtr

                While True
                    destinationhWnd = GetWindow(sourceformhWnd, GetWindowType.GW_HWNDNEXT)

                    If (destinationhWnd = IntPtr.Zero) Then
                        Return False
                    End If

                    If (destinationhWnd = sourceformhWnd) Then
                        Return False
                    End If

                    Dim windowStyle As Integer = GetWindowStyle(destinationhWnd)
                    If ((windowStyle And WindowStyles.WS_VISIBLE) = 0) Then
                        Continue While
                    End If

                    If (GetWindowTextLength(destinationhWnd) = 0) Then
                        Continue While
                    End If

                    ' Skip the taskbar window (required for Windows 7,
                    ' when the target window Is the only other window
                    ' in the taskbar)
                    If (IsTaskBar(destinationhWnd)) Then
                        Continue While
                    End If

                    Exit While
                End While

                Return EnsureForegroundWindow(destinationhWnd)

            Catch ex As Exception
                Return False
            End Try
        End Function

        Private Function IsTaskBar(ByVal hWnd As IntPtr) As Boolean
            Dim p As Process = Nothing

            Try
                Dim strText As String = GetWindowText(hWnd, True)
                If (strText = Nothing) Then
                    Return False
                End If
                If (Not strText.Equals("Start", StringComparison.OrdinalIgnoreCase)) Then
                    Return False
                End If
                If (Not strText.Equals("Inicio", StringComparison.OrdinalIgnoreCase)) Then
                    Return False
                End If

                Dim ProcessId As IntPtr
                GetWindowThreadProcessId(hWnd, ProcessId)
                If ProcessId = IntPtr.Zero Then
                    Return False
                End If

                p = Process.GetProcessById(ProcessId.ToInt32)
                Dim strExe As String = CardonerSistemas.Files.GetFileNameFromFullPath(p.MainModule.FileName).Trim()

                Return strExe.Equals("explorer.exe", StringComparison.OrdinalIgnoreCase)

            Catch ex As Exception
            Finally
                Try
                    If (Not p Is Nothing) Then
                        p.Dispose()
                    End If
                Catch ex As Exception
                End Try
            End Try

            Return False
        End Function

        Private Function EnsureForegroundWindow(ByVal hWnd As IntPtr) As Boolean
            If Not IsWindow(hWnd) Then
                Return False
            End If

            Dim hWndInit As IntPtr = GetForegroundWindow()

            If Not SetForegroundWindow(hWnd) Then
                Return False
            End If

            Dim nStartMS As Integer = Environment.TickCount
            While (Environment.TickCount - nStartMS) < 1000
                Dim h As IntPtr = GetForegroundWindow()
                If (h = hWnd) Then
                    Return True
                End If

                ' Some applications (Like Microsoft Edge) have multiple
                ' windows And automatically redirect the focus to other
                ' windows, thus also break when a different window gets
                ' focused (except when h Is zero, which can occur while
                ' the focus transfer occurs)
                If (h <> IntPtr.Zero) AndAlso (h <> hWndInit) Then
                    Return True
                End If

                Application.DoEvents()
            End While

            Return False
        End Function

#End Region

    End Module

End Namespace
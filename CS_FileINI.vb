Imports System.Runtime.InteropServices
Imports System.Text

Module CS_FileINI
    Private Declare Auto Function GetPrivateProfileSectionNames Lib "kernel32" (ByVal lpReturnedString As StringBuilder, ByVal nSize As Integer, ByVal lpFileName As String) As Integer
    Private Declare Function GetPrivateProfileSection Lib "kernel32" Alias "GetPrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpReturnedString As IntPtr, ByVal nSize As Integer, ByVal lpFileName As String) As Integer
    '<DllImport("kernel32.dll", SetLastError:=True)>
    'Private Function GetPrivateProfileSection(ByVal lpAppName As String, ByVal lpReturnedString As IntPtr, ByVal nSize As Integer, ByVal lpFileName As String) As Integer

    'End Function


    Private Declare Auto Function GetPrivateProfileString Lib "kernel32" (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As StringBuilder, ByVal nSize As Integer, ByVal lpFileName As String) As Integer

    Private Declare Auto Function WritePrivateProfileString Lib "Kernel32" (ByVal IpApplication As String, ByVal Ipkeyname As String, ByVal IpString As String, ByVal IpFileName As String) As Integer

    Public Enum CSINIDataTypes
        csidtString
        csidtNumberInteger
        csidtNumberDecimal
        csidtDateTime
        csidtBoolean
    End Enum

    Public Function GetSectionsNames(ByVal FileName As String, Optional ByVal FilterPrefix As String = "") As Collection
        Dim sbValue As StringBuilder
        Dim Value As String
        Dim ReturnedValue As Integer

        Dim aValues() As String
        Dim Index As Long

        sbValue = New StringBuilder(100000000)
        Value = ""
        GetSectionsNames = New Collection

        ReturnedValue = GetPrivateProfileSectionNames(sbValue, sbValue.Capacity, FileName)
        If ReturnedValue > 0 Then
            Value = Left(Value, ReturnedValue)
            If Right(Value, 1) = vbNullChar Then
                Value = Left(Value, ReturnedValue - 1)
            End If
            aValues = Split(Value, vbNullChar)
            On Error Resume Next
            For Index = 0 To UBound(aValues)
                If Left(aValues(Index), Len(FilterPrefix)) = FilterPrefix Then
                    GetSectionsNames.Add(aValues(Index), KEY_STRINGER & aValues(Index))
                End If
            Next Index
        End If
    End Function

    Public Function GetSectionsNames_FromApplication(Optional ByVal FilterPrefix As String = "") As Collection
        GetSectionsNames_FromApplication = GetSectionsNames(My.Application.Info.DirectoryPath & "\" & My.Application.Info.AssemblyName & ".ini", FilterPrefix)
    End Function

    Public Function GetSectionValues(ByVal FileName As String, ByVal Section As String) As System.Collections.Specialized.NameValueCollection
        Const MaxIniBuffer As Integer = 100000000

        Dim pBuffer As IntPtr
        Dim bytesRead As Integer
        Dim sectionData As New System.Text.StringBuilder(MaxIniBuffer)

        Dim values As New System.Collections.Specialized.NameValueCollection
        Dim pos As Integer ' seperator position
        Dim name, value As String
        GetSectionValues = New System.Collections.Specialized.NameValueCollection

        'Extracted from http://www.pinvoke.net/default.aspx/kernel32/GetPrivateProfileSection.html
        'get a pointer to the unmanaged memory
        pBuffer = Marshal.AllocHGlobal(MaxIniBuffer * Marshal.SizeOf(New Char()))

        bytesRead = GetPrivateProfileSection(Section, pBuffer, MaxIniBuffer, FileName)
        If bytesRead > 0 Then
            For i As Integer = 0 To bytesRead - 1
                sectionData.Append(Convert.ToChar(Marshal.ReadByte(pBuffer, i)))
            Next

            sectionData.Remove(sectionData.Length - 1, 1)

            For Each line As String In sectionData.ToString().Split(Convert.ToChar(0))
                ' locate the seperator
                pos = line.IndexOf("=")

                If (pos > -1) Then

                    ' get values
                    name = line.Substring(0, pos)
                    value = line.Substring(pos + 1)

                    ' add to collection
                    values.Add(name, value)
                End If
            Next
        Else
            values = Nothing
        End If

        ' release the unmanaged memory
        Marshal.FreeHGlobal(pBuffer)

        ' return collection or Nothing if we weren't able to get anything
        Return values
    End Function

    Public Function GetValue(ByVal FileName As String, ByVal Section As String, ByVal Key As String, ByVal DefaultValue As Object, ByVal DataType As CSINIDataTypes) As Object
        Dim sbValue As StringBuilder
        Dim Value As String
        Dim ReturnedValue As Integer

        sbValue = New StringBuilder(500)
        Value = ""
        ReturnedValue = GetPrivateProfileString(Section, Key, "", sbValue, sbValue.Capacity, FileName)
        If ReturnedValue > 0 Then
            Value = sbValue.ToString
        End If

        On Error Resume Next

        Select Case DataType
            Case CSINIDataTypes.csidtString
                If ReturnedValue > 0 Then
                    GetValue = CStr(Value)
                Else
                    GetValue = CStr(DefaultValue)
                End If

            Case CSINIDataTypes.csidtNumberInteger
                If ReturnedValue > 0 And IsNumeric(Value) Then
                    GetValue = CLng(Value)
                Else
                    GetValue = CLng(DefaultValue)
                End If

            Case CSINIDataTypes.csidtNumberDecimal
                If ReturnedValue > 0 And IsNumeric(Value) Then
                    GetValue = CDec(Value)
                Else
                    GetValue = CDec(DefaultValue)
                End If

            Case CSINIDataTypes.csidtDateTime
                If ReturnedValue > 0 And IsDate(Value) Then
                    GetValue = CDate(Value)
                Else
                    GetValue = CDate(DefaultValue)
                End If

            Case CSINIDataTypes.csidtBoolean
                If ReturnedValue > 0 Then
                    If IsNumeric(Value) Then
                        GetValue = CBool(Value <> 0)
                    ElseIf Value = "False" Or Value = "Falso" Then
                        GetValue = False
                    ElseIf Value = "True" Or Value = "Verdadero" Then
                        GetValue = True
                    Else
                        GetValue = CBool(DefaultValue)
                    End If
                Else
                    GetValue = CBool(DefaultValue)
                End If
        End Select
    End Function

    Public Function GetValueAsString(ByVal FileName As String, ByVal Section As String, ByVal Key As String, ByVal DefaultValue As String) As String
        Return GetValue(FileName, Section, Key, DefaultValue, CSINIDataTypes.csidtString)
    End Function

    Public Function GetValueAsInteger(ByVal FileName As String, ByVal Section As String, ByVal Key As String, ByVal DefaultValue As Integer) As Integer
        Return GetValue(FileName, Section, Key, DefaultValue, CSINIDataTypes.csidtNumberInteger)
    End Function

    Public Function GetValueAsDecimal(ByVal FileName As String, ByVal Section As String, ByVal Key As String, ByVal DefaultValue As Decimal) As Decimal
        Return GetValue(FileName, Section, Key, DefaultValue, CSINIDataTypes.csidtNumberDecimal)
    End Function

    Public Function GetValueAsDateTime(ByVal FileName As String, ByVal Section As String, ByVal Key As String, ByVal DefaultValue As Date) As Date
        Return GetValue(FileName, Section, Key, DefaultValue, CSINIDataTypes.csidtDateTime)
    End Function

    Public Function GetValueAsBoolean(ByVal FileName As String, ByVal Section As String, ByVal Key As String, ByVal DefaultValue As Boolean) As Boolean
        Return GetValue(FileName, Section, Key, DefaultValue, CSINIDataTypes.csidtBoolean)
    End Function

    Public Function GetValue_FromApplication(ByVal Section As String, ByVal Key As String, ByVal DefaultValue As Object, ByVal DataType As CSINIDataTypes) As Object
        GetValue_FromApplication = GetValue(My.Application.Info.DirectoryPath & "\" & My.Application.Info.AssemblyName & ".ini", Section, Key, DefaultValue, DataType)
    End Function

    Public Function SetValue(ByVal FileName As String, ByVal Section As String, ByVal Key As String, ByVal Value As String) As Boolean
        SetValue = (WritePrivateProfileString(Section, Key, Value, FileName) = 0)
    End Function

    Public Function SetValue_ToApplication(ByVal Section As String, ByVal Key As String, ByVal Value As String) As Boolean
        SetValue_ToApplication = SetValue(My.Application.Info.DirectoryPath & "\" & My.Application.Info.AssemblyName & ".ini", Section, Key, Value)
    End Function
End Module
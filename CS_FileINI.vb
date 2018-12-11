Option Strict Off
'
' IniFile class 
' by Todd Davis (toddhd@hotmail.com)
'
' Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated 
' documentation files (the "Software"), to deal in the Software without restriction, including without limitation 
' the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and 
' to permit persons to whom the Software is furnished to do so, subject to the following conditions:
'
' This permission notice shall be included in all copies or substantial portions 
' of the Software.
'
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED 
' TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL 
' THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF 
' CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER 
' DEALINGS IN THE SOFTWARE.
'

Friend Class IniFile
    Private Sections As New ArrayList
    Private Const COMMENTSTRING As String = ";"

    Friend Function Read(ByVal Filename As String, Optional ByVal CreateIfNotExist As Boolean = True) As Boolean
        If Not System.IO.File.Exists(Filename) Then ' verify that the file exists
            If CreateIfNotExist Then 'if it does not exist, check to see if we should create it
                Try
                    Dim fs As System.IO.FileStream = System.IO.File.Create(Filename) 'create it
                    fs.Close()
                Catch ex As Exception
                    MsgBox("Error: Cannot create file " & Filename & vbCrLf & ex.ToString(), MsgBoxStyle.Critical, My.Application.Info.Title)
                    Return False
                End Try
            Else
                MsgBox("Error: File " & Filename & " does not exist. Cannot create IniFile.", MsgBoxStyle.Critical, My.Application.Info.Title)
                Return False
            End If
        End If

        Sections.Clear() 'Clear the arraylist

        Try
            Dim sr As New System.IO.StreamReader(Filename) 'Declare a streamreader
            Dim CurrentSection As String = "" 'Flag to track what section we are currently in
            Dim ThisLine As String = sr.ReadLine() 'Read in the first line

            Do
                Select Case Eval(ThisLine) 'Evalue the contents of the line
                    Case "Section"
                        AddSection(RemoveBrackets(RemoveComment(ThisLine)), IsCommented(ThisLine)) 'Add the section to the sections arraylist
                        CurrentSection = RemoveBrackets(RemoveComment(ThisLine)) 'Make this the current section, so we know where to keys to
                    Case "Key"
                        AddKey(GetKeyName(ThisLine), GetKeyValue(ThisLine), CurrentSection, IsCommented(ThisLine))
                    Case "Comment"
                        'AddComment(ThisLine, CurrentSection)
                    Case "Blank"
                        'TODO: Should we create a blank object to handle blanks?
                    Case ""
                        'We hit something unknown - ignore it
                End Select
                ThisLine = sr.ReadLine() 'Get the next line
            Loop Until ThisLine Is Nothing 'continue until the end of the file
            sr.Close() 'close the file

            Return True

        Catch ex As Exception
            MsgBox("Error: " & ex.ToString, MsgBoxStyle.Critical, My.Application.Info.Title)
            Return False
        End Try

    End Function

    Private Function Eval(ByVal value As String) As String
        value = Trim(value) 'Remove any leading/trailing spaces, just in case they exist

        'If the value is blank, then it is a blank line
        If value = "" Then Return "Blank"

        'If the value is surrounded by brackets, then it is a section
        If Microsoft.VisualBasic.Left(RemoveComment(value), 1) = "[" And Microsoft.VisualBasic.Right(value, 1) = "]" Then Return "Section"

        'If the value contains an equals sign (=), then it is a value. This test can be fooled by 
        'comment with an equals sign in it, but it is the best test we have. We test for this before
        'testing for a comment in case the key is commented out. It is still a key.
        If InStr(value, "=", CompareMethod.Text) > 0 Then Return "Key"

        'If the value is preceeded by the comment string, then it is a pure comment
        If IsCommented(value) Then Return "Comment"

        Return ""
    End Function

    Private Function GetKeyName(ByVal Value As String) As String
        'If the value is commented out, then remove the comment string so we can get the name
        If IsCommented(Value) Then Value = RemoveComment(Value)
        Dim Equals As Integer = InStr(Value, "=", CompareMethod.Text) 'Locate the equals sign
        If Equals > 0 Then 'It should be, but just to be safe
            Return Microsoft.VisualBasic.Left(Value, Equals - 1) 'Return everything before the equals sign
        Else : Return ""
        End If
    End Function

    Private Function GetKeyValue(ByVal value As String) As String
        Dim Equals As Integer = InStr(value, "=", CompareMethod.Text) 'Locate the equals sign
        If Equals > 0 Then 'It should be, but just to be safe
            Return Microsoft.VisualBasic.Right(value, Len(value) - Equals) 'Return everything after the equals sign
        Else : Return ""
        End If
    End Function

    Private Function IsCommented(ByVal value As String) As Boolean
        'Return true if the passed value starts with a comment string
        If Microsoft.VisualBasic.Left(value, Len(CommentString)) = CommentString Then Return True
        Return False
    End Function

    Private Function RemoveComment(ByVal value As String) As String
        'Return the value with the comment string stripped
        Return CStr(IIf(IsCommented(value), value.Remove(0, Len(CommentString)), value))
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Adds a key/value to a given section. If the section does not exist, it is created.
    ''' </summary>
    ''' <param name="KeyName">The name of the key to add. If the key alreadys exists, then no action is taken.</param>
    ''' <param name="KeyValue">The value to assign to the new key.</param>
    ''' <param name="SectionName">The section to add the new key to. If it does not exist, it is created.</param>
    ''' <param name="IsCommented">Optional, defaults to false. Will create the key in commented state.</param>
    ''' <param name="InsertBefore">Optional. Will insert the new key prior to the specified key.</param>
    ''' <returns></returns>
    ''' <remarks>If the section does not exist, it will be created. If the 'IsCommented' option is true, then the newly created section will also be commented. If the 'InsertBefore' option is used, the specified key does not exist, then the new key is simply added to the section. If the section the key is being added to is commented, then the key will be commented as well.
    ''' </remarks>
    ''' <history>
    ''' 	[TDavis]	1/19/2004	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Function AddKey(ByVal KeyName As String, ByVal KeyValue As String, ByVal SectionName As String, Optional ByVal IsCommented As Boolean = False, Optional ByVal InsertBefore As String = Nothing) As Boolean
        Dim ThisSection As Section = GetSection(SectionName) 'verify that the section exists

        If ThisSection Is Nothing Then
            AddSection(SectionName, IsCommented)
        End If
        If ThisSection.IsCommented Then  'If the section is commented out, then this key must be too
            IsCommented = True
        End If
        If Not GetKey(KeyName, SectionName) Is Nothing Then  'verify that the key does *not* exist
            Return False
        End If

        Dim ThisKey As New Key(KeyName, KeyValue, IsCommented) 'create a new key

        If InsertBefore Is Nothing Then 'if no insertbefore is required
            ThisSection.Add(ThisKey) 'then add the new key to the bottom of the section
            Return True
        Else
            Dim KeyIndex = GetKeyIndex(InsertBefore, SectionName) 'locate the key to insert prior to
            If KeyIndex > -1 Then 'if the key exists
                ThisSection.Insert(KeyIndex, ThisKey) 'then do the insert
                Return True
            Else
                ThisSection.Add(ThisKey) 'the key to insert prior to wasn't found, so just add it
                Return False 'the key to insert prior to was not found
            End If
        End If
    End Function

    Private Function GetKeyIndex(ByVal KeyName As String, ByVal SectionName As String) As Integer
        'returns the index of a given key
        'Dim ThisKey As Key = GetKey(KeyName, SectionName)
        'If ThisKey Is Nothing Then Return -1
        'Dim ThisSection As Section = GetSection(SectionName) 
        'Return ThisSection.IndexOf(ThisKey.Name)

        Dim SectionEnumerator As System.Collections.IEnumerator = Sections.GetEnumerator()
        While SectionEnumerator.MoveNext()
            If SectionEnumerator.Current.Name = SectionName Then
                Dim KeyEnumerator As System.Collections.IEnumerator = SectionEnumerator.Current.GetEnumerator()
                While KeyEnumerator.MoveNext()
                    If KeyEnumerator.Current.Name = KeyName Then Return SectionEnumerator.Current.indexof(KeyEnumerator.Current)
                End While
            End If
        End While
        Return -1
    End Function
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Adds a section to the IniFile. If the section already exists, then no action is taken.
    ''' </summary>
    ''' <param name="SectionName">The name of the section to add.</param>
    ''' <param name="IsCommented">Optional, defaults to false. Will add the section in a commented state.</param>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[TDavis]	1/19/2004	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Sub AddSection(ByVal SectionName As String, Optional ByVal IsCommented As Boolean = False)
        If GetSection(SectionName) Is Nothing Then Sections.Add(New Section(SectionName, IsCommented)) 'Add the section to the sections arraylist
    End Sub

    Private Function GetSection(ByVal SectionName As String) As Section
        'Return the given section object
        Dim myEnumerator As System.Collections.IEnumerator = Sections.GetEnumerator()
        While myEnumerator.MoveNext()
            Dim CurrentSection As Section = myEnumerator.Current
            If LCase(CurrentSection.Name) = LCase(SectionName) Then Return myEnumerator.Current
        End While
        Return Nothing
    End Function
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Return the sections in the IniFile.
    ''' </summary>
    ''' <returns>Returns an ArrayList of Section objects.</returns>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[TDavis]	1/19/2004	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Function GetSections() As ArrayList
        'returns an arraylist of the sections in the inifile
        Dim ListOfSections As New ArrayList
        Dim myEnumerator As System.Collections.IEnumerator = Sections.GetEnumerator()
        While myEnumerator.MoveNext()
            ListOfSections.Add(myEnumerator.Current)
        End While
        Return ListOfSections
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Returns an arraylist of Key objects in a given Section. Section must exist.
    ''' </summary>
    ''' <param name="SectionName">The name of the Section to retrieve the keys from.</param>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[TDavis]	1/19/2004	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Function GetKeys(ByVal SectionName As String) As ArrayList
        'returns an arraylist of the keys in a given section
        Dim ListOfKeys As New ArrayList
        Dim ThisSection As Section = GetSection(SectionName)
        If ThisSection Is Nothing Then Return Nothing
        Dim KeyEnumerator As System.Collections.IEnumerator = ThisSection.GetEnumerator()
        While KeyEnumerator.MoveNext()
            ListOfKeys.Add(KeyEnumerator.Current)
        End While

        Return ListOfKeys
    End Function
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Comments a given section, including all of the keys contained in the section.
    ''' </summary>
    ''' <param name="SectionName">The name of the Section to comment.</param>
    ''' <returns></returns>
    ''' <remarks>Keys that are already commented will <b>not</b> preserve their comment status if 'UnCommentSection' is used later on.
    ''' </remarks>
    ''' <history>
    ''' 	[TDavis]	1/19/2004	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Function CommentSection(ByVal SectionName As String) As Boolean
        'Comments a given section and all of its keys
        Dim ThisSection As Section = GetSection(SectionName)
        If ThisSection Is Nothing Then Return False
        ThisSection.IsCommented = True
        Dim myEnumerator As System.Collections.IEnumerator = ThisSection.GetEnumerator()
        While myEnumerator.MoveNext()
            myEnumerator.Current.IsCommented = True
        End While
        Return True
    End Function
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Uncomments a given section, and all of its keys.
    ''' </summary>
    ''' <param name="SectionName">The name of the Section to uncomment.</param>
    ''' <returns></returns>
    ''' <remarks>Any keys in the section that were previously commented will be uncommented after this function.
    ''' </remarks>
    ''' <history>
    ''' 	[TDavis]	1/19/2004	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Function UnCommentSection(ByVal SectionName As String) As Boolean
        'Uncomments a given section and all of its keys
        Dim ThisSection As Section = GetSection(SectionName)
        If ThisSection Is Nothing Then Return False
        ThisSection.IsCommented = False
        Dim myEnumerator As System.Collections.IEnumerator = ThisSection.GetEnumerator()
        While myEnumerator.MoveNext()
            myEnumerator.Current.IsCommented = False
        End While
        Return True
    End Function
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Comments a given key in a given section. Both the key and the section must exist. 
    ''' </summary>
    ''' <param name="KeyName">The name of the key to comment.</param>
    ''' <param name="SectionName">The name of the section the key is in.</param>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[TDavis]	1/19/2004	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Function CommentKey(ByVal KeyName As String, ByVal SectionName As String) As Boolean
        'Comments a given a key
        Dim ThisKey As Key = GetKey(KeyName, SectionName)
        If ThisKey Is Nothing Then
            Return False
        Else
            Return True
        End If

        ThisKey.IsCommented = True
    End Function
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Uncomments a given key in a given section. Both the key and section must exist.
    ''' </summary>
    ''' <param name="KeyName">The name of the key to uncomment.</param>
    ''' <param name="SectionName">The name of the section the key is in.</param>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[TDavis]	1/19/2004	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Function UnCommentKey(ByVal KeyName As String, ByVal SectionName As String) As Boolean
        'Uncomments a given key
        Dim ThisKey As Key = GetKey(KeyName, SectionName)
        If ThisKey Is Nothing Then
            Return False
        Else
            Return True
        End If
        ThisKey.IsCommented = False
    End Function
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Renames a section. The section must exist.
    ''' </summary>
    ''' <param name="SectionName">The name of the section to be renamed.</param>
    ''' <param name="NewSectionName">The new name of the section.</param>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[TDavis]	1/19/2004	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Function RenameSection(ByVal SectionName As String, ByVal NewSectionName As String) As Boolean
        Dim ThisSection As Section = GetSection(SectionName)
        If ThisSection Is Nothing Then Return False
        ThisSection.Name = NewSectionName
        Return True
    End Function
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Renames a given key key in a given section. Both they key and the section must exist. The value is not altered.
    ''' </summary>
    ''' <param name="KeyName">The name of the key to be renamed.</param>
    ''' <param name="SectionName">The name of the section the key exists in.</param>
    ''' <param name="NewKeyName">The new name of the key.</param>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[TDavis]	1/19/2004	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Function RenameKey(ByVal KeyName As String, ByVal SectionName As String, ByVal NewKeyName As String) As Boolean
        Dim ThisKey As Key = GetKey(KeyName, SectionName)
        If ThisKey Is Nothing Then Return False
        ThisKey.Name = NewKeyName
        Return True
    End Function
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Changes the value of a given key in a given section. Both the key and the section must exist.
    ''' </summary>
    ''' <param name="KeyName">The name of the key whose value should be changed.</param>
    ''' <param name="SectionName">The name of the section the key exists in.</param>
    ''' <param name="NewValue">The new value to assign to the key.</param>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[TDavis]	1/19/2004	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Function ChangeValue(ByVal KeyName As String, ByVal SectionName As String, ByVal NewValue As String) As Boolean
        Dim ThisSection As Section = GetSection(SectionName)
        If ThisSection Is Nothing Then Return False
        Dim ThisKey As Key = GetKey(KeyName, SectionName)
        If ThisKey Is Nothing Then Return False
        ThisKey.Value = NewValue
        Return True
    End Function

    Friend Function GetKey(ByVal KeyName As String, ByRef SectionObj As Section) As Key
        If SectionObj Is Nothing Then Return Nothing
        Dim myEnumerator As System.Collections.IEnumerator = SectionObj.GetEnumerator()
        While myEnumerator.MoveNext()
            If LCase(myEnumerator.Current.Name) = LCase(KeyName) Then Return myEnumerator.Current
        End While
        Return Nothing
    End Function

    Friend Function GetKey(ByVal KeyName As String, ByVal SectionName As String) As Key
        Dim ThisSection As Section = GetSection(SectionName)
        If ThisSection Is Nothing Then
            Return Nothing
        Else
            Return GetKey(KeyName, ThisSection)
        End If
    End Function
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Deletes a given section. The section must exist. All the keys in the section will also be deleted.
    ''' </summary>
    ''' <param name="SectionName">The name of the section to be deleted.</param>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[TDavis]	1/19/2004	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Function DeleteSection(ByVal SectionName As String) As Boolean
        Dim ThisSection As Section = GetSection(SectionName)
        If ThisSection Is Nothing Then Return False
        Sections.Remove(ThisSection)
        Return True
    End Function
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Deletes a given key in a given section. Both the key and the section must exist.
    ''' </summary>
    ''' <param name="KeyName">The name of the key to be deleted.</param>
    ''' <param name="SectionName">The name of the section the key exists in.</param>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[TDavis]	1/19/2004	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Function DeleteKey(ByVal KeyName As String, ByVal SectionName As String) As Boolean
        Dim ThisSection As Section = GetSection(SectionName)
        If ThisSection Is Nothing Then Return False
        Dim ThisKey As Key = GetKey(KeyName, SectionName)
        If ThisKey Is Nothing Then Return False
        ThisSection.Remove(ThisKey)
        Return True
    End Function

    Public Sub Save(ByVal FileName As String)
        If System.IO.File.Exists(FileName) Then System.IO.File.Delete(FileName) ' Remove the existing file

        'Loop through the arraylist (Content) and write each line to the file
        Dim sw As New System.IO.StreamWriter(FileName)

        Dim SectionEnumerator As System.Collections.IEnumerator = Sections.GetEnumerator()
        While SectionEnumerator.MoveNext()
            sw.Write(AddBrackets(SectionEnumerator.Current.Name) & vbCrLf)
            Dim KeyEnumerator As System.Collections.IEnumerator = SectionEnumerator.Current.GetEnumerator()
            While KeyEnumerator.MoveNext()
                sw.Write(KeyEnumerator.Current.Name & "=" & KeyEnumerator.Current.Value & vbCrLf)
            End While
        End While
        sw.Close()
    End Sub

    Public Sub SaveXML(ByVal FileName As String, Optional ByVal Encode As System.Text.Encoding = Nothing)
        Dim strXMLPath As String = FileName
        Dim objXMLWriter As System.Xml.XmlTextWriter
        objXMLWriter = New System.Xml.XmlTextWriter(strXMLPath, Encode) 'Create a new XML file

        objXMLWriter.WriteStartDocument()
        objXMLWriter.WriteStartElement("configuration")

        Dim SectionEnumerator As System.Collections.IEnumerator = Sections.GetEnumerator()
        While SectionEnumerator.MoveNext()
            objXMLWriter.WriteStartElement("section")
            objXMLWriter.WriteAttributeString("name", SectionEnumerator.Current.Name)
            Dim KeyEnumerator As System.Collections.IEnumerator = SectionEnumerator.Current.GetEnumerator()
            While KeyEnumerator.MoveNext()
                objXMLWriter.WriteStartElement("setting")
                objXMLWriter.WriteAttributeString("name", KeyEnumerator.Current.Name)
                objXMLWriter.WriteAttributeString("value", KeyEnumerator.Current.Value)
                objXMLWriter.WriteEndElement()
            End While
            objXMLWriter.WriteEndElement()
        End While

        objXMLWriter.WriteEndElement() 'write the ending tag for configuration
        objXMLWriter.WriteEndDocument()
        objXMLWriter.Flush()
        objXMLWriter.Close()
    End Sub

    Private Function GetText(Optional ByVal ReturnAsHTML As Boolean = False) As String
        Dim CrLf As String = IIf(ReturnAsHTML, "<br>", vbCrLf)
        Dim sb As New System.Text.StringBuilder
        Dim SectionEnumerator As System.Collections.IEnumerator = Sections.GetEnumerator()
        While SectionEnumerator.MoveNext()
            sb.Append(AddBrackets(SectionEnumerator.Current.Name) & CrLf)
            Dim KeyEnumerator As System.Collections.IEnumerator = SectionEnumerator.Current.GetEnumerator()
            While KeyEnumerator.MoveNext()
                sb.Append(KeyEnumerator.Current.Name & "=" & KeyEnumerator.Current.Value & CrLf)
            End While
        End While
        Return sb.ToString
    End Function

    Private Function RemoveBrackets(ByVal Value As String) As String
        Dim chArr() As Char = {"[", "]", " "}
        Value = Value.TrimStart(chArr)
        Value = Value.TrimEnd(chArr)
        Return Value
    End Function

    Private Function AddBrackets(ByVal Value As String) As String
        Return "[" & Trim(Value) & "]"
    End Function

End Class

Public Class Section
    Inherits ArrayList
    Public Name As String
    Public IsCommented As Boolean

    Public Sub New(ByVal SectionName As String, Optional ByVal SectionIsCommented As Boolean = False)
        Name = SectionName
        IsCommented = SectionIsCommented
    End Sub

    Public Overloads Overrides Function Equals(ByVal obj As Object) As Boolean
        If obj Is Nothing Or Not Me.GetType() Is obj.GetType() Then
            Return False
        End If
        Dim s As Section = CType(obj, Section)
        Return Me.Name = s.Name And Me.IsCommented = s.IsCommented
    End Function

    Public Overrides Function GetHashCode() As Integer
        Dim s As String = Name
        Return s.GetHashCode()
    End Function
End Class

Public Class Key
    Public Name As String
    Public Value As String
    Public IsCommented As Boolean

    Public Sub New(ByVal KeyName As String, ByVal KeyValue As String, Optional ByVal KeyIsCommented As Boolean = False)
        Name = KeyName
        Value = KeyValue
        IsCommented = KeyIsCommented
    End Sub

    Public Overloads Overrides Function Equals(ByVal obj As Object) As Boolean
        If obj Is Nothing Or Not Me.GetType() Is obj.GetType() Then
            Return False
        End If
        Dim k As Key = CType(obj, Key)
        Return Me.Name = k.Name And Me.Value = k.Value And Me.IsCommented = k.IsCommented
    End Function

    Public Overrides Function GetHashCode() As Integer
        Dim s As String = Name + Value + IsCommented
        Return s.GetHashCode()
    End Function
End Class



'Imports System.Text
'Imports System.Runtime.InteropServices

'Module CS_FileINI
'    Private Declare Auto Function GetPrivateProfileSectionNames Lib "kernel32" (ByVal lpReturnedString As StringBuilder, ByVal nSize As Integer, ByVal lpFileName As String) As Integer
'    Private Declare Function GetPrivateProfileSection Lib "kernel32" Alias "GetPrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpReturnedString As IntPtr, ByVal nSize As Integer, ByVal lpFileName As String) As Integer
'    '<DllImport("kernel32.dll", SetLastError:=True)>
'    'Private Function GetPrivateProfileSection(ByVal lpAppName As String, ByVal lpReturnedString As IntPtr, ByVal nSize As Integer, ByVal lpFileName As String) As Integer

'    'End Function

'    Private Declare Auto Function GetPrivateProfileString Lib "kernel32" (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As StringBuilder, ByVal nSize As Integer, ByVal lpFileName As String) As Integer

'    Private Declare Auto Function WritePrivateProfileString Lib "Kernel32" (ByVal IpApplication As String, ByVal Ipkeyname As String, ByVal IpString As String, ByVal IpFileName As String) As Integer

'    Friend Enum CSINIDataTypes
'        csidtString
'        csidtNumberInteger
'        csidtNumberDecimal
'        csidtDateTime
'        csidtBoolean
'    End Enum

'    Friend Function GetSectionsNamesAsArray(ByVal FileName As String, Optional ByVal FilterPrefix As String = "") As String()
'        Dim sbValue As StringBuilder
'        Dim Value As String
'        Dim ReturnedValue As Integer

'        Dim aValues() As String

'        sbValue = New StringBuilder(100000000)

'        ReturnedValue = GetPrivateProfileSectionNames(sbValue, sbValue.Capacity, FileName)
'        If ReturnedValue > 0 Then
'            Value = Left(Value, ReturnedValue)
'            If Right(Value, 1) = vbNullChar Then
'                Value = Left(Value, ReturnedValue - 1)
'            End If
'            aValues = Split(Value, vbNullChar)
'        End If

'        Return aValues
'    End Function


'    Public Function GetSectionsNamesAsCollection(ByVal FileName As String, Optional ByVal FilterPrefix As String = "") As Collection
'        Dim aValues() As String
'        Dim Index As Integer
'        Dim CollectionToReturn As New Collection

'        aValues = GetSectionsNamesAsArray(FileName, FilterPrefix)

'        On Error Resume Next
'        For Index = 0 To UBound(aValues)
'            If Left(aValues(Index), Len(FilterPrefix)) = FilterPrefix Then
'                CollectionToReturn.Add(aValues(Index), aValues(Index))
'            End If
'        Next Index
'        Return CollectionToReturn
'    End Function

'    Public Function GetSectionsNames_FromApplication(Optional ByVal FilterPrefix As String = "") As Collection
'        GetSectionsNames_FromApplication = GetSectionsNamesAsCollection(My.Application.Info.DirectoryPath & "\" & My.Application.Info.AssemblyName & ".ini", FilterPrefix)
'    End Function

'    Public Function GetSectionValues(ByVal FileName As String, ByVal Section As String) As System.Collections.Specialized.NameValueCollection
'        Const MaxIniBuffer As Integer = 100000000

'        Dim pBuffer As IntPtr
'        Dim bytesRead As Integer
'        Dim sectionData As New System.Text.StringBuilder(MaxIniBuffer)

'        Dim values As New System.Collections.Specialized.NameValueCollection
'        Dim pos As Integer ' seperator position
'        Dim name, value As String
'        GetSectionValues = New System.Collections.Specialized.NameValueCollection

'        'Extracted from http://www.pinvoke.net/default.aspx/kernel32/GetPrivateProfileSection.html
'        'get a pointer to the unmanaged memory
'        pBuffer = Marshal.AllocHGlobal(MaxIniBuffer * Marshal.SizeOf(New Char()))

'        bytesRead = GetPrivateProfileSection(Section, pBuffer, MaxIniBuffer, FileName)
'        If bytesRead > 0 Then
'            For i As Integer = 0 To bytesRead - 1
'                sectionData.Append(Convert.ToChar(Marshal.ReadByte(pBuffer, i)))
'            Next

'            sectionData.Remove(sectionData.Length - 1, 1)

'            For Each line As String In sectionData.ToString().Split(Convert.ToChar(0))
'                ' locate the seperator
'                pos = line.IndexOf("=")

'                If (pos > -1) Then

'                    ' get values
'                    name = line.Substring(0, pos)
'                    value = line.Substring(pos + 1)

'                    ' add to collection
'                    values.Add(name, value)
'                End If
'            Next
'        Else
'            values = Nothing
'        End If

'        ' release the unmanaged memory
'        Marshal.FreeHGlobal(pBuffer)

'        ' return collection or Nothing if we weren't able to get anything
'        Return values
'    End Function

'    Public Function GetValue(ByVal FileName As String, ByVal Section As String, ByVal Key As String, ByVal DefaultValue As Object, ByVal DataType As CSINIDataTypes) As Object
'        Dim sbValue As StringBuilder
'        Dim Value As String
'        Dim ReturnedValue As Integer

'        sbValue = New StringBuilder(500)
'        Value = ""
'        ReturnedValue = GetPrivateProfileString(Section, Key, "", sbValue, sbValue.Capacity, FileName)
'        If ReturnedValue > 0 Then
'            Value = sbValue.ToString
'        End If

'        On Error Resume Next

'        Select Case DataType
'            Case CSINIDataTypes.csidtString
'                If ReturnedValue > 0 Then
'                    GetValue = CStr(Value)
'                Else
'                    GetValue = CStr(DefaultValue)
'                End If

'            Case CSINIDataTypes.csidtNumberInteger
'                If ReturnedValue > 0 And IsNumeric(Value) Then
'                    GetValue = CLng(Value)
'                Else
'                    GetValue = CLng(DefaultValue)
'                End If

'            Case CSINIDataTypes.csidtNumberDecimal
'                If ReturnedValue > 0 And IsNumeric(Value) Then
'                    GetValue = CDec(Value)
'                Else
'                    GetValue = CDec(DefaultValue)
'                End If

'            Case CSINIDataTypes.csidtDateTime
'                If ReturnedValue > 0 And IsDate(Value) Then
'                    GetValue = CDate(Value)
'                Else
'                    GetValue = CDate(DefaultValue)
'                End If

'            Case CSINIDataTypes.csidtBoolean
'                If ReturnedValue > 0 Then
'                    If IsNumeric(Value) Then
'                        GetValue = CBool(CInt(Value) <> 0)
'                    ElseIf Value = "False" Or Value = "Falso" Then
'                        GetValue = False
'                    ElseIf Value = "True" Or Value = "Verdadero" Then
'                        GetValue = True
'                    Else
'                        GetValue = CBool(DefaultValue)
'                    End If
'                Else
'                    GetValue = CBool(DefaultValue)
'                End If
'        End Select
'    End Function

'    Public Function GetValueAsString(ByVal FileName As String, ByVal Section As String, ByVal Key As String, ByVal DefaultValue As String) As String
'        Return CStr(GetValue(FileName, Section, Key, DefaultValue, CSINIDataTypes.csidtString))
'    End Function

'    Public Function GetValueAsInteger(ByVal FileName As String, ByVal Section As String, ByVal Key As String, ByVal DefaultValue As Integer) As Integer
'        Return CInt(GetValue(FileName, Section, Key, DefaultValue, CSINIDataTypes.csidtNumberInteger))
'    End Function

'    Public Function GetValueAsDecimal(ByVal FileName As String, ByVal Section As String, ByVal Key As String, ByVal DefaultValue As Decimal) As Decimal
'        Return CInt(GetValue(FileName, Section, Key, DefaultValue, CSINIDataTypes.csidtNumberDecimal))
'    End Function

'    Public Function GetValueAsDateTime(ByVal FileName As String, ByVal Section As String, ByVal Key As String, ByVal DefaultValue As Date) As Date
'        Return CDate(GetValue(FileName, Section, Key, DefaultValue, CSINIDataTypes.csidtDateTime))
'    End Function

'    Public Function GetValueAsBoolean(ByVal FileName As String, ByVal Section As String, ByVal Key As String, ByVal DefaultValue As Boolean) As Boolean
'        Return CBool(GetValue(FileName, Section, Key, DefaultValue, CSINIDataTypes.csidtBoolean))
'    End Function

'    Public Function GetValue_FromApplication(ByVal Section As String, ByVal Key As String, ByVal DefaultValue As Object, ByVal DataType As CSINIDataTypes) As Object
'        GetValue_FromApplication = GetValue(My.Application.Info.DirectoryPath & "\" & My.Application.Info.AssemblyName & ".ini", Section, Key, DefaultValue, DataType)
'    End Function

'    Public Function SetValue(ByVal FileName As String, ByVal Section As String, ByVal Key As String, ByVal Value As String) As Boolean
'        SetValue = (WritePrivateProfileString(Section, Key, Value, FileName) = 0)
'    End Function

'    Public Function SetValue_ToApplication(ByVal Section As String, ByVal Key As String, ByVal Value As String) As Boolean
'        SetValue_ToApplication = SetValue(My.Application.Info.DirectoryPath & "\" & My.Application.Info.AssemblyName & ".ini", Section, Key, Value)
'    End Function
'End Module
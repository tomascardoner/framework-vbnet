Imports System.IO
Imports System.Reflection

Module CS_Icon

    Friend Function GetIconFromEmbeddedResource(ByVal IconName As String, ByVal IconSize As Size) As Icon
        Dim asm As Assembly
        Dim rnames As String()
        Dim tofind As String

        asm = Assembly.GetExecutingAssembly()
        rnames = asm.GetManifestResourceNames()
        tofind = "." + IconName + ".ICO"

        For Each rname As String In rnames
            If rname.EndsWith(tofind, StringComparison.CurrentCultureIgnoreCase) Then
                Using IconStream As Stream = asm.GetManifestResourceStream(rname)
                    Return New Icon(IconStream, IconSize)
                End Using
            End If
        Next

        Throw New ArgumentException("Icon not found")
    End Function

End Module
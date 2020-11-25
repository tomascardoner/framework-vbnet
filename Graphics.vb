Namespace CardonerSistemas

    Module Graphics

        Friend Function GetIconFromBitmap(ByRef bitmap As Bitmap) As Icon
            Dim pointerIcon As IntPtr = bitmap.GetHicon()
            Using icon As Icon = Icon.FromHandle(pointerIcon)
                Return icon
            End Using
        End Function

    End Module
End Namespace
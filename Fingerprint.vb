Namespace CardonerSistemas

    Friend Module Fingerprint
        Friend Enum AnsiFingerIndexValues As Byte
            RightThumb = 1
            RightIndex = 2
            RightMiddle = 3
            RightRing = 4
            RightLittle = 5
            LeftThumb = 6
            LeftIndex = 7
            LeftMiddle = 8
            LeftRing = 9
            LeftLittle = 10
        End Enum

        Friend Enum EnrollmentMaskValues As Integer
            LeftThumb = 1
            LeftIndex = 2
            LeftMiddle = 4
            LeftRing = 8
            LeftLittle = 16
            RightThumb = 32
            RightIndex = 64
            RightMiddle = 128
            RightRing = 256
            RightLittle = 512
        End Enum

        Friend Function GetAnsiFingerIndexValue(value As Byte) As AnsiFingerIndexValues
            Return CType(value, AnsiFingerIndexValues)
        End Function

        Friend Function GetEnrollmentMaskValue(fingerIndexValue As AnsiFingerIndexValues) As EnrollmentMaskValues
            Dim ansiFingerIndexValueName As String
            Dim enrollmentMaskValue As EnrollmentMaskValues

            ansiFingerIndexValueName = [Enum].GetName(GetType(AnsiFingerIndexValues), fingerIndexValue)
            enrollmentMaskValue = DirectCast([Enum].Parse(GetType(EnrollmentMaskValues), ansiFingerIndexValueName), EnrollmentMaskValues)

            Return enrollmentMaskValue
        End Function

        Friend Function GetEnrollmentMaskValue(value As Byte) As EnrollmentMaskValues
            Return GetEnrollmentMaskValue(GetAnsiFingerIndexValue(value))
        End Function
    End Module

End Namespace
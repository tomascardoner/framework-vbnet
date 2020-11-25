Namespace CardonerSistemas

    Module DateTime

        Friend Function GetElapsedCompleteYearsFromDates(ByVal startDate As Date, ByVal endDate As Date) As Long
            Dim elapsedYears As Long

            elapsedYears = DateAndTime.DateDiff(DateInterval.Year, startDate, endDate)
            If (startDate.Month > endDate.Month) Or (startDate.Month = endDate.Month AndAlso startDate.Day > endDate.Day) Then
                elapsedYears -= 1
            End If

            Return elapsedYears
        End Function

        Friend Function GetElapsedYearsMonthsAndDaysFromDays(ByVal elapsedDays As Integer) As String
            Dim daysInAYear As Decimal
            Dim elapsedYears As Short
            Dim elapsedMonths As Byte
            Dim result As String

            If elapsedDays > 1460 Then
                ' Elapsed more than 4 years, so take aproximate account of leap years
                daysInAYear = CDec(365.25)
            Else
                daysInAYear = 365
            End If

            ' Get elapsed years and the remaining days
            elapsedYears = CShort(elapsedDays / daysInAYear)
            elapsedDays = CInt(elapsedDays Mod daysInAYear)

            ' Get elapsed months and the remainig days
            elapsedMonths = CByte(elapsedDays / 30)
            elapsedDays = elapsedDays Mod 30

            ' Years
            Select Case elapsedYears
                Case 0
                    result = ""
                Case 1
                    result = "1 año"
                Case Else
                    result = elapsedYears & " años"
            End Select

            ' Months
            If elapsedMonths > 0 Then
                If result <> "" Then
                    If elapsedDays = 0 Then
                        result &= " y "
                    Else
                        result &= ", "
                    End If
                End If

                If elapsedMonths = 1 Then
                    result &= "1 mes"
                Else
                    result &= elapsedMonths & " meses"
                End If
            End If

            ' Days
            If elapsedDays > 0 Then
                If result <> "" Then
                    result &= " y "
                End If

                If elapsedDays = 1 Then
                    result &= "1 día"
                Else
                    result &= elapsedDays & " días"
                End If
            End If

            ' Final dot
            If result <> "" Then
                result &= "."
            End If

            Return result
        End Function

    End Module
    
End Namespace
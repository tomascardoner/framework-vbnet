Namespace CardonerSistemas

    Friend Module DateTime

#Region "Declarations"

        Friend Enum PeriodTypes As Byte
            All
            Day
            Week
            Month
            Year
            Range
        End Enum

        Friend Enum PeriodDayValues As Byte
            DayToday
            DayYesterday
            DayBeforeYesterday
            DayLast2
            DayLast3
            DayLast4
            DayLast7
            DayLast15
        End Enum

        Friend Enum PeriodWeekValues As Byte
            WeekCurrent
            WeekBeforeCurrent
            WeekLast2
            WeekLast3
        End Enum

        Friend Enum PeriodMonthValues As Byte
            MonthCurrent
            MonthBeforeCurrent
            MonthLast2
            MonthLast3
            MonthLast6
        End Enum

        Friend Enum PeriodYearValues As Byte
            YearCurrent
            YearBeforeCurrent
        End Enum

        Friend Enum PeriodRangeValues As Byte
            DateEqual
            DateBefore
            DateAfter
            DateBetween
        End Enum

#End Region

#Region "ComboBoxs de filtros de fechas"

        Friend Sub FillPeriodTypesComboBox(ByRef control As Windows.Forms.ComboBox, selectedPeriodType As PeriodTypes)
            control.Items.AddRange({My.Resources.STRING_ITEM_ALL_MALE, "Día:", "Semana:", "Mes:", "Año:", "Fecha"})
            control.SelectedIndex = CInt(selectedPeriodType)
        End Sub

        Friend Sub FillPeriodValuesComboBox(ByRef control As Windows.Forms.ComboBox, periodType As PeriodTypes)
            control.Items.Clear()
            Select Case periodType
                Case PeriodTypes.All
                    control.Items.Add(String.Empty)
                Case PeriodTypes.Day
                    control.Items.AddRange({"Hoy", "Ayer", "Anteayer", "Últimos 2", "Últimos 3", "Últimos 4", "Últimos 7", "Últimos 15"})
                Case PeriodTypes.Week
                    control.Items.AddRange({"Actual", "Anterior", "Últimas 2", "Últimas 3"})
                Case PeriodTypes.Month
                    control.Items.AddRange({"Actual", "Anterior", "Últimos 2", "Últimos 3", "Últimos 6"})
                Case PeriodTypes.Year
                    control.Items.AddRange({"Actual", "Anterior"})
                Case PeriodTypes.Range
                    control.Items.AddRange({"es igual a:", "es anterior a:", "es posterior a:", "está entre:"})
            End Select
            control.Visible = (periodType <> PeriodTypes.All)
            control.SelectedIndex = 0
        End Sub

        Friend Sub GetDatesFromPeriodTypeAndValue(periodType As PeriodTypes, periodValue As Byte, ByRef dateFrom As Date, ByRef dateTo As Date, dateValueFrom As Date, dateValueTo As Date)
            Select Case periodType
                ' All
                Case PeriodTypes.All
                    dateFrom = Date.MinValue
                    dateTo = Date.MaxValue

                ' Days
                Case PeriodTypes.Day
                    Dim periodDayValue As PeriodDayValues = CType(periodValue, PeriodDayValues)

                    Select Case periodDayValue
                        Case PeriodDayValues.DayToday
                            dateFrom = System.DateTime.Today
                            dateTo = dateFrom.AddDays(1).AddMilliseconds(-1)
                        Case PeriodDayValues.DayYesterday
                            dateFrom = System.DateTime.Today.AddDays(-1)
                            dateTo = dateFrom.AddDays(1).AddMilliseconds(-1)
                        Case PeriodDayValues.DayBeforeYesterday
                            dateFrom = System.DateTime.Today.AddDays(-2)
                            dateTo = dateFrom.AddDays(1).AddMilliseconds(-1)
                        Case PeriodDayValues.DayLast2
                            dateFrom = System.DateTime.Today.AddDays(-1)
                            dateTo = System.DateTime.Today.AddDays(1).AddMilliseconds(-1)
                        Case PeriodDayValues.DayLast3
                            dateFrom = System.DateTime.Today.AddDays(-2)
                            dateTo = System.DateTime.Today.AddDays(1).AddMilliseconds(-1)
                        Case PeriodDayValues.DayLast4
                            dateFrom = System.DateTime.Today.AddDays(-3)
                            dateTo = System.DateTime.Today.AddDays(1).AddMilliseconds(-1)
                        Case PeriodDayValues.DayLast7
                            dateFrom = System.DateTime.Today.AddDays(-6)
                            dateTo = System.DateTime.Today.AddDays(1).AddMilliseconds(-1)
                        Case PeriodDayValues.DayLast15
                            dateFrom = System.DateTime.Today.AddDays(-14)
                            dateTo = System.DateTime.Today.AddDays(1).AddMilliseconds(-1)
                    End Select

                ' Weeks
                Case PeriodTypes.Week
                    Dim periodWeekValue As PeriodWeekValues = CType(periodValue, PeriodWeekValues)

                    Select Case periodWeekValue
                        Case PeriodWeekValues.WeekCurrent
                            dateFrom = System.DateTime.Today.AddDays(-System.DateTime.Today.DayOfWeek)
                            dateTo = System.DateTime.Today.AddDays(1).AddMilliseconds(-1)
                        Case PeriodWeekValues.WeekBeforeCurrent
                            dateFrom = System.DateTime.Today.AddDays(-System.DateTime.Today.DayOfWeek - 7)
                            dateTo = System.DateTime.Today.AddDays(-System.DateTime.Today.DayOfWeek).AddMilliseconds(-1)
                        Case PeriodWeekValues.WeekLast2
                            dateFrom = System.DateTime.Today.AddDays(-System.DateTime.Today.DayOfWeek - 7)
                            dateTo = System.DateTime.Today.AddDays(7 - System.DateTime.Today.DayOfWeek).AddMilliseconds(-1)
                        Case PeriodWeekValues.WeekLast3
                            dateFrom = System.DateTime.Today.AddDays(-System.DateTime.Today.DayOfWeek - 14)
                            dateTo = System.DateTime.Today.AddDays(7 - System.DateTime.Today.DayOfWeek).AddMilliseconds(-1)
                    End Select

                ' Months
                Case PeriodTypes.Month
                    Dim periodMonthValue As PeriodMonthValues = CType(periodValue, PeriodMonthValues)

                    Select Case periodMonthValue
                        Case PeriodMonthValues.MonthCurrent
                            dateFrom = New Date(System.DateTime.Today.Year, System.DateTime.Today.Month, 1, 0, 0, 0, DateTimeKind.Local)
                            dateTo = New Date(System.DateTime.Today.Year, System.DateTime.Today.Month, 1, 0, 0, 0, DateTimeKind.Local).AddMonths(1).AddMilliseconds(-1)
                        Case PeriodMonthValues.MonthBeforeCurrent
                            dateFrom = New Date(System.DateTime.Today.Year, System.DateTime.Today.Month, 1, 0, 0, 0, DateTimeKind.Local).AddMonths(-1)
                            dateTo = New Date(System.DateTime.Today.Year, System.DateTime.Today.Month, 1, 0, 0, 0, DateTimeKind.Local).AddMilliseconds(-1)
                        Case PeriodMonthValues.MonthLast2
                            dateFrom = New Date(System.DateTime.Today.Year, System.DateTime.Today.Month, 1, 0, 0, 0, DateTimeKind.Local).AddMonths(-1)
                            dateTo = New Date(System.DateTime.Today.Year, System.DateTime.Today.Month, 1, 0, 0, 0, DateTimeKind.Local).AddMonths(1).AddMilliseconds(-1)
                        Case PeriodMonthValues.MonthLast3
                            dateFrom = New Date(System.DateTime.Today.Year, System.DateTime.Today.Month, 1, 0, 0, 0, DateTimeKind.Local).AddMonths(-2)
                            dateTo = New Date(System.DateTime.Today.Year, System.DateTime.Today.Month, 1, 0, 0, 0, DateTimeKind.Local).AddMonths(1).AddMilliseconds(-1)
                        Case PeriodMonthValues.MonthLast6
                            dateFrom = New Date(System.DateTime.Today.Year, System.DateTime.Today.Month, 1, 0, 0, 0, DateTimeKind.Local).AddMonths(-5)
                            dateTo = New Date(System.DateTime.Today.Year, System.DateTime.Today.Month, 1, 0, 0, 0, DateTimeKind.Local).AddMonths(1).AddMilliseconds(-1)
                    End Select

                ' Years
                Case PeriodTypes.Year
                    Dim periodYearValue As PeriodYearValues = CType(periodValue, PeriodYearValues)

                    Select Case periodYearValue
                        Case PeriodYearValues.YearCurrent
                            dateFrom = New Date(System.DateTime.Today.Year, 1, 1, 0, 0, 0, DateTimeKind.Local)
                            dateTo = New Date(System.DateTime.Today.Year, 1, 1, 0, 0, 0, DateTimeKind.Local).AddYears(1).AddMilliseconds(-1)
                        Case PeriodYearValues.YearBeforeCurrent
                            dateFrom = New Date(System.DateTime.Today.Year, 1, 1, 0, 0, 0, DateTimeKind.Local).AddYears(-1)
                            dateTo = New Date(System.DateTime.Today.Year, 1, 1, 0, 0, 0, DateTimeKind.Local).AddMilliseconds(-1)
                    End Select

                ' Range
                Case PeriodTypes.Range
                    Dim periodRangeValue As PeriodRangeValues = CType(periodValue, PeriodRangeValues)

                    Select Case periodRangeValue
                        Case PeriodRangeValues.DateEqual
                            dateFrom = dateValueFrom
                            dateTo = dateValueFrom.AddDays(1).AddMilliseconds(-1)
                        Case PeriodRangeValues.DateBefore
                            dateFrom = Date.MinValue
                            dateTo = dateValueFrom.AddMilliseconds(-1)
                        Case PeriodRangeValues.DateAfter
                            dateFrom = dateValueFrom.AddDays(1)
                            dateTo = Date.MaxValue
                        Case PeriodRangeValues.DateBetween
                            dateFrom = dateValueFrom
                            dateTo = dateValueTo.AddDays(1).AddMilliseconds(-1)
                    End Select
            End Select
        End Sub

#End Region


        Friend Function GetElapsedCompleteYearsFromDates(startDate As Date, endDate As Date) As Long
            Dim elapsedYears As Long

            elapsedYears = DateAndTime.DateDiff(DateInterval.Year, startDate, endDate)
            If (startDate.Month > endDate.Month) OrElse (startDate.Month = endDate.Month AndAlso startDate.Day > endDate.Day) Then
                elapsedYears -= 1
            End If

            Return elapsedYears
        End Function

        Friend Function GetElapsedYearsMonthsAndDaysFromDays(elapsedDays As Integer) As String
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
                    result = $"{elapsedYears} años"
            End Select

            ' Months
            If elapsedMonths > 0 Then
                If result.Length > 0 Then
                    If elapsedDays = 0 Then
                        result &= " y "
                    Else
                        result &= ", "
                    End If
                End If

                If elapsedMonths = 1 Then
                    result &= "1 mes"
                Else
                    result &= $"{elapsedMonths} meses"
                End If
            End If

            ' Days
            If elapsedDays > 0 Then
                If result.Length > 0 Then
                    result &= " y "
                End If

                If elapsedDays = 1 Then
                    result &= "1 día"
                Else
                    result &= $"{elapsedDays} días"
                End If
            End If

            ' Final dot
            If result.Length > 0 Then
                result &= "."
            End If

            Return result
        End Function

    End Module

End Namespace
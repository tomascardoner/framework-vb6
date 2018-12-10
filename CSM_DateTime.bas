Attribute VB_Name = "CSM_DateTime"
Option Explicit

'/////////////////////////////////////////
'FIRST DATE OF MONTH
Public Function GetFirstDateOfMonth(ByVal Month As Byte, ByVal Year As Integer) As Date
    GetFirstDateOfMonth = DateSerial(Year, Month, 1)
End Function

Public Function GetFirstDateOfMonthFromDate(ByVal DateValue As Date) As Date
    GetFirstDateOfMonthFromDate = DateSerial(Year(DateValue), Month(DateValue), 1)
End Function

Public Function GetFirstDateOfPreviousMonthFromDate(ByVal DateValue As Date) As Date
    GetFirstDateOfPreviousMonthFromDate = GetFirstDateOfMonthFromDate(DateAdd("m", -1, DateValue))
End Function

'/////////////////////////////////////////
'FIRST DATE OF MONTH
Public Function GetLastDayOfMonth(ByVal Month As Byte, ByVal Year As Integer) As Byte
    'SIMPLEST METHOD
    GetLastDayOfMonth = Day(DateAdd("d", -1, DateAdd("m", 1, DateSerial(Year, Month, 1))))
    
'    'MANUAL METHOD
'    Select Case Month
'        Case 1, 3, 5, 7, 8, 10, 12
'            'Enero, Marzo, Mayo, Julio, Agosto, Octubre, Diciembre
'            GetLastDayOfMonth = 31
'        Case 4, 6, 9, 11
'            'Abril, Junio, Septiembre, Noviembre
'            GetLastDayOfMonth = 30
'        Case 2
'            'Febrero...calcular bisiestos
'            If IsLeapYear(Year) Then
'                GetLastDayOfMonth = 29
'            Else
'                GetLastDayOfMonth = 28
'            End If
'    End Select
End Function

Public Function GetLastDateOfMonth(ByVal Month As Byte, ByVal Year As Integer) As Byte
    GetLastDateOfMonth = DateAdd("d", -1, DateAdd("m", 1, DateSerial(Year, Month, 1)))
End Function

Public Function GetLastDateOfMonthFromDate(ByVal DateValue As Date) As Date
    GetLastDateOfMonthFromDate = DateAdd("d", -1, DateAdd("m", 1, DateSerial(Year(DateValue), Month(DateValue), 1)))
End Function

Public Function GetLastDateOfPreviousMonthFromDate(ByVal DateValue As Date) As Date
    GetLastDateOfPreviousMonthFromDate = GetLastDateOfMonthFromDate(DateAdd("m", -1, DateValue))
End Function

Public Function IsLeapYear(ByVal Year As Integer) As Boolean
    'SIMPLEST METHOD
    IsLeapYear = (Day(DateAdd("d", -1, DateSerial(Year, 3, 1))) = 29)

'    'MANUAL METHOD
'    If (Year Mod 4) = 0 Then
'        'Es divisible por 4
'        If (Year Mod 100) = 0 Then
'            'Es divisible por 100
'            If (Year Mod 400) = 0 Then
'                'Es divisible por 400
'                IsLeapYear = True
'            Else
'                'No es divisible por 400
'                IsLeapYear = False
'            End If
'        Else
'            'No es divisible por 100
'            IsLeapYear = True
'        End If
'    Else
'        'No es divisible por 4
'        IsLeapYear = False
'    End If
End Function

Public Function GetElapsedDayString(ByVal dtFechaInicio As Date, ByVal dtFechaFin As Date) As String
    Dim lngYearsElapsedTemp As Long
    Dim dtCompleteYear As Date
    Dim lngEdadAnios As Long
    Dim strYearsString As String
    
    Dim lngMonthsElapsedTemp As Long
    Dim dtCompleteMonth As Date
    Dim lngEdadMeses As Long
    Dim strMonthsString As String
    
    Dim lngEdadDias As Long
    Dim strDaysString As String
    
    '=========================== AÑOS ==================================================
    'Calculo la cantidad de años transcurridos
    lngYearsElapsedTemp = DateDiff("yyyy", dtFechaInicio, dtFechaFin)
    
    'A la fecha de hoy le resto la cantidad de años que calculé anteriormente,
    'esto lo hago para corregir el error del VB
    dtCompleteYear = DateAdd("yyyy", -lngYearsElapsedTemp, dtFechaFin)
    
    'Si me pasé del límite, le resto un año
    If dtCompleteYear < dtFechaInicio Then
        dtCompleteYear = DateAdd("yyyy", 1, dtCompleteYear)
    End If
    
    'Calculo los años reales
    lngEdadAnios = DateDiff("yyyy", dtCompleteYear, dtFechaFin)
    
    Select Case lngEdadAnios
        Case 0
        Case 1
            strYearsString = "1 año"
        Case Else
            strYearsString = Format(lngEdadAnios) + " años"
    End Select
    '===================================================================================
    
    '=========================== MESES =================================================
    'Calculo la cantidad de meses transcurridos desde el último Año Completo
    lngMonthsElapsedTemp = DateDiff("m", dtFechaInicio, dtCompleteYear)
    
    'A la fecha de hoy le resto la cantidad de años que calculé anteriormente,
    'esto lo hago para corregir el error del VB
    dtCompleteMonth = DateAdd("m", -lngMonthsElapsedTemp, dtCompleteYear)
    
    'Si me pasé del límite, le resto un año
    If dtCompleteMonth < dtFechaInicio Then
        dtCompleteMonth = DateAdd("m", 1, dtCompleteMonth)
    End If
       
    'Calculo los años reales
    lngEdadMeses = DateDiff("m", dtCompleteMonth, dtCompleteYear)
    
    Select Case lngEdadMeses
        Case 0
        Case 1
            strMonthsString = "1 mes"
        Case Else
            strMonthsString = Format(lngEdadMeses) + " meses"
    End Select
    '===================================================================================
    
    '=========================== DIAS ==================================================
    'Calculo los días restantes
    lngEdadDias = DateDiff("d", dtFechaInicio, dtCompleteMonth)
    
    Select Case lngEdadDias
        Case 0
        Case 1
            strDaysString = "1 día"
        Case Else
            strDaysString = Format(lngEdadDias) + " días"
    End Select
    '===================================================================================
    
    'Armo el string final
    If strYearsString <> "" And strMonthsString <> "" And strDaysString <> "" Then
        GetElapsedDayString = strYearsString + ", " + strMonthsString + " y " + strDaysString
    Else
        If strYearsString <> "" And strMonthsString <> "" Then
            GetElapsedDayString = strYearsString + " y " + strMonthsString
        Else
            If strYearsString <> "" And strDaysString <> "" Then
                GetElapsedDayString = strYearsString + " y " + strDaysString
            Else
                If strMonthsString <> "" And strDaysString <> "" Then
                    GetElapsedDayString = strMonthsString + " y " + strDaysString
                Else
                    GetElapsedDayString = strYearsString + strMonthsString + strDaysString
                End If
            End If
        End If
    End If
End Function

Public Sub GetWeekDates(ByVal WeekNumber As Byte, ByVal YearNumber As Integer, ByRef FirstDate As Date, ByRef LastDate As Date)
    Dim Jan1st As Long
    Dim FirstSunday As Long
    
    Jan1st = DateSerial(YearNumber, 1, 1)  ' get 1/1/yyyy
    FirstSunday = Jan1st - Weekday(Jan1st) + 1 ' get the sunday in week 1
  
    'Check if 01/01/yyyy is in week 1
    If CInt(Format(Jan1st, "ww")) <> 1 Then
        FirstSunday = FirstSunday + 7
    End If
  
    FirstDate = FirstSunday + 7 * (WeekNumber - 1)
    LastDate = FirstDate + 6
End Sub

Public Function WeekNumber(ByVal dDate As Date, Optional ByVal FirstDayOfWeek As VbDayOfWeek = vbUseSystemDayOfWeek) As Integer
    WeekNumber = CInt(Format(dDate, "ww", FirstDayOfWeek))
    
'    Dim d2 As Date
'
'    d2 = DateSerial(Year(dDate - Weekday(dDate - 1) + 4), 1, 3)
'    WeekNumber = Int((dDate - d2 + Weekday(d2) + 5) / 7)
End Function

Public Function YearStart(ByVal iWhichYear As Integer, Optional ByVal FirstDayOfWeek As VbDayOfWeek = vbUseSystemDayOfWeek) As Date
    Dim iWeekDay As VbDayOfWeek
    Dim dNewYear As Date
    
    dNewYear = DateSerial(iWhichYear, 1, 1)

    iWeekDay = Weekday(dNewYear, FirstDayOfWeek)
    
    YearStart = DateAdd("d", (FirstDayOfWeek - iWeekDay), dNewYear)
    If iWeekDay < FirstDayOfWeek Then
        YearStart = DateAdd("d", 7, dNewYear)
    End If
End Function

Public Function WeeksInYear(ByVal iYear As Integer) As Integer
    WeeksInYear = WeekNumber(DateAdd("d", -1, YearStart(iYear + 1)))
End Function

Public Function WeekStart(ByVal iYear As Integer, ByVal iWeek As Integer) As Date
    WeekStart = DateAdd("ww", iWeek - 1, YearStart(iYear))
End Function

Public Function WeekEnd(ByVal iYear As Integer, ByVal iWeek As Integer) As Date
    WeekEnd = DateAdd("d", 6, DateAdd("ww", iWeek - 1, YearStart(iYear)))
End Function

Public Function WeekStart_NotRegular(ByVal iYear As Integer, ByVal iWeek As Integer, ByVal FirstDayOfWeek As VbDayOfWeek) As Date
    Dim WorkDate As Date
    
    WorkDate = DateAdd("ww", iWeek - 1, YearStart(iYear))
    
    WeekStart_NotRegular = DateAdd("d", FirstDayOfWeek - Weekday(WorkDate), WorkDate)
End Function

Public Function WeekEnd_NotRegular(ByVal iYear As Integer, ByVal iWeek As Integer, ByVal FirstDayOfWeek As VbDayOfWeek) As Date
    Dim WorkDate As Date
    
    WorkDate = DateAdd("ww", iWeek - 1, YearStart(iYear))
    
    WorkDate = DateAdd("d", 6, WorkDate)
    
    WeekEnd_NotRegular = DateAdd("d", FirstDayOfWeek - Weekday(WorkDate), WorkDate)
End Function

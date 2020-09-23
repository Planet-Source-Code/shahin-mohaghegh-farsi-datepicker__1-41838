Attribute VB_Name = "modDate"
Dim NY, NLY, LY, ALY As Integer

Public Function sh2mi(myd As String) As Variant
Dim FiDay As String

If myd = "  /  /  " Then
    sh2mi = Null
End If
posd1 = InStr(1, myd, "/")
posd2 = InStr(posd1 + 1, myd, "/")
'posd1 = At("/", myd, 1)
'posd2 = At("/", myd, 2)
yy = Val(VBA.Left(myd, posd1 - 1))
mm = Val(Mid(myd, posd1 + 1, posd2 - posd1))
dd = Val(VBA.Right(myd, Len(myd) - posd2))
If yy = 0 Or mm = 0 Or dd = 0 Then
    sh2mi = Null
    Exit Function
End If
'***** Check to see if all characters are Numeric First For DD
If mm > 6 And dd > 30 Then
    sh2mi = Null
End If
If yy = 0 Or (mm < 1 Or mm > 12) Or (dd < 1 Or dd > 31) Then
    sh2mi = Null
End If
'****** Otherwise Everything is ok and continue
If yy > 1000 Then
    ayy = yy + 621
Else
    ayy = yy + 1921
End If
If yy < 75 Or yy Mod 4 <> 3 Then
    leap = 1
    leapd = -1
Else
    leap = 0
    leapd = 0
End If
'*** now for month and day
Select Case mm
    Case 1
        If dd <= 12 + leapd Then
            amm = 3
            Add = dd + 19 + leap
        Else
            amm = 4
            Add = dd - 12 + leap
        End If
    Case 2
        If dd <= 11 + leapd Then
            amm = 4
            Add = dd + 19 + leap
        Else
            amm = 5
            Add = dd - 11 + leap
        End If
    Case 3
        If dd <= 11 + leapd Then
            amm = 5
            Add = dd + 20 + leap
        Else
            amm = 6
            Add = dd - 11 + leap
        End If
    Case 4
        If dd <= 10 + leapd Then
            amm = 6
            Add = dd + 20 + leap
        Else
            amm = 7
            Add = dd - 10 + leap
        End If
    Case 5
        If dd <= 10 + leapd Then
            amm = 7
            Add = dd + 21 + leap
        Else
            amm = 8
            Add = dd - 10 + leap
        End If
    Case 6
        If dd <= 10 + leapd Then
            amm = 8
            Add = dd + 21 + leap
        Else
            amm = 9
            Add = dd - 10 + leap
        End If
    Case 7
        If dd <= 9 + leapd Then
            amm = 9
            Add = dd + 21 + leap
        Else
            amm = 10
            Add = dd - 9 + leap
        End If
    Case 8
        If dd <= 10 + leapd Then
            amm = 10
            Add = dd + 21 + leap
        Else
            amm = 11
            Add = dd - 10 + leap
        End If
    Case 9
        If dd <= 10 + leapd Then
            amm = 11
            Add = dd + 20 + leap
        Else
            amm = 12
            Add = dd - 10 + leap
        End If
    Case 10
        If dd <= 11 + leapd Then
            amm = 12
            Add = dd + 20 + leap
        Else
            amm = 1
            ayy = ayy + 1
            Add = dd - 11 + leap
        End If
    Case 11
        If dd <= 12 + leapd Then
            amm = 1
            ayy = ayy + 1
            Add = dd + 19 + leap
        Else
            amm = 2
            ayy = ayy + 1
            Add = dd - 12 + leap
        End If
    Case 12
        If dd <= 10 + leapd And ayy + 1 Mod 4 <> 0 Then
            amm = 2
            ayy = ayy + 1
            Add = dd + 18 + leap
        Else
            If dd <= 11 + leapd And ayy + 1 Mod 4 = 0 Then
                amm = 2
                ayy = ayy + 1
                Add = dd + 18 + leap
            Else
                If dd > 10 + leapd And ayy + 1 Mod 4 <> 0 Then
                    amm = 3
                    ayy = ayy + 1
                    Add = dd - 10 + leap
                Else
                    If dd > 11 + leapd And ayy + 1 Mod 4 = 0 Then
                        amm = 3
                        ayy = ayy + 1
                        Add = dd - 11 + leap
                    End If
                End If
            End If
        End If
End Select
FiDay = CStr(ayy) + "/" + CStr(amm) + "/" + CStr(Add)
FiDay = CDate(FiDay)
sh2mi = FiDay
End Function

Public Function mi2sh(MiDate As Date, Optional TwoDigits As Boolean)
If TypeName(MiDate) <> "Date" Then
        mi2sh = "  /  /  "
        Exit Function
End If
MiYear = Year(MiDate)
NY = 0: NLY = 0: LY = 0: ALY = 0
Select Case MiYear
    Case 0
        mi2sh = "  /  /  "
        Exit Function
    Case 0 To 1996
        If MiYear Mod 4 > 0 Then
            NY = 1
        Else
            If MiYear Mod 4 = 0 Then
                NLY = 1
            End If
        End If
    Case 1996 To 9999
        If MiYear Mod 4 = 0 Then
            LY = 1
        Else
            If MiYear Mod 4 = 1 Then
                ALY = 1
            Else
                If MiYear Mod 4 > 1 Then
                    NY = 1
                End If
            End If
        End If
End Select
mi2sh = CnvrtDate(MiDate, TwoDigits)
End Function

Private Function CnvrtDate(MiDate As Date, Optional TwoDigits As Boolean)
Dim ShYear As Integer
MiMoon = Month(MiDate)
MiDay = Day(MiDate)
ShYear = Year(MiDate) - IIf(MiMoon > 3 Or (MiMoon = 3 And MiDay > 20 - LY), 621, 622)
Select Case MiMoon
    Case 1
        ShMoon = IIf(MiDay > 20 - ALY, 11, 10)
        ShDay = IIf(MiDay > 20 - ALY, MiDay - 20 + ALY, MiDay + 10 + ALY)
    Case 2
        ShMoon = IIf(MiDay > 19 - ALY, 12, 11)
        ShDay = IIf(MiDay > 19 - ALY, MiDay - 19 + ALY, MiDay + 11 + ALY)
    Case 3
        ShMoon = IIf(MiDay > 20 - LY, 1, 12)
        ShDay = IIf(MiDay > 20 - LY, MiDay - 20 + LY, MiDay + 10 - NY)
    Case 4
        ShMoon = IIf(MiDay > 20 - LY, 2, 1)
        ShDay = IIf(MiDay > 20 - LY, MiDay - 20 + LY, MiDay + 11 + LY)
    Case 5
        ShMoon = IIf(MiDay > 21 - LY, 3, 2)
        ShDay = IIf(MiDay > 21 - LY, MiDay - 21 + LY, MiDay + 10 + LY)
    Case 6
        ShMoon = IIf(MiDay > 21 - LY, 4, 3)
        ShDay = IIf(MiDay > 21 - LY, MiDay - 21 + LY, MiDay + 10 + LY)
    Case 7
        ShMoon = IIf(MiDay > 22 - LY, 5, 4)
        ShDay = IIf(MiDay > 22 - LY, MiDay - 22 + LY, MiDay + 9 + LY)
    Case 8
        ShMoon = IIf(MiDay > 22 - LY, 6, 5)
        ShDay = IIf(MiDay > 22 - LY, MiDay - 22 + LY, MiDay + 9 + LY)
    Case 9
        ShMoon = IIf(MiDay > 22 - LY, 7, 6)
        ShDay = IIf(MiDay > 22 - LY, MiDay - 22 + LY, MiDay + 9 + LY)
    Case 10
        ShMoon = IIf(MiDay > 22 - LY, 8, 7)
        ShDay = IIf(MiDay > 22 - LY, MiDay - 22 + LY, MiDay + 8 + LY)
    Case 11
        ShMoon = IIf(MiDay > 21 - LY, 9, 8)
        ShDay = IIf(MiDay > 21 - LY, MiDay - 21 + LY, MiDay + 9 + LY)
    Case 12
        ShMoon = IIf(MiDay > 21 - LY, 10, 9)
        ShDay = IIf(MiDay > 21 - LY, MiDay - 21 + LY, MiDay + 9 + LY)
End Select
If TwoDigits Then

   CnvrtDate = VBA.Right(CStr(ShYear), 2) + "/" + VBA.Right(CStr(100 + ShMoon), 2) + "/" + VBA.Right(CStr(100 + ShDay), 2)
Else
'   mi2sh = cstr(ShYear, 4) + "/" + vba.Right(Str(100 + ShMoon, 3), 2) + "/" + vba.Right(Str(100 + ShDay, 3), 2)
   CnvrtDate = CStr(ShYear) + "/" + VBA.Right(CStr(100 + ShMoon), 2) + "/" + VBA.Right(CStr(100 + ShDay), 2)
End If
End Function

Public Function WeekString(TheDay As Integer)
Select Case TheDay
    Case vbSaturday
        WeekString = "Saturday"
    Case vbSunday
        WeekString = "Sunday"
    Case vbMonday
        WeekString = "Monday"
    Case vbTuesday
        WeekString = "Tuesday"
    Case vbWednesday
        WeekString = "Wednesday"
    Case vbThursday
        WeekString = "Thursday"
    Case vbFriday
        WeekString = "Friday"
End Select

End Function

Public Function TWeekDay(TheDay As Date) As Integer
Select Case Weekday(TheDay)
    Case vbSaturday
        TWeekDay = 1
    Case vbSunday
        TWeekDay = 2
    Case vbMonday
        TWeekDay = 3
    Case vbTuesday
        TWeekDay = 4
    Case vbWednesday
        TWeekDay = 5
    Case vbThursday
        TWeekDay = 6
    Case vbFriday
        TWeekDay = 7
End Select
End Function

Public Function FWeekString(TheDay As Integer)
Select Case TheDay
    Case vbSaturday
        FWeekString = "‘‰»Â"
    Case vbSunday
        FWeekString = "Ìﬂ‘‰»Â"
    Case vbMonday
        FWeekString = "œÊ‘‰»Â"
    Case vbTuesday
        FWeekString = "”Â ‘‰»Â"
    Case vbWednesday
        FWeekString = "çÂ«—‘‰»Â"
    Case vbThursday
        FWeekString = "Å‰Ã‘‰»Â"
    Case vbFriday
        FWeekString = "Ã„⁄Â"
End Select

End Function

Public Function MMonthName(Mnth As Integer)
Select Case Mnth
    Case 1
        MMonthName = "January"
    Case 2
        MMonthName = "February"
    Case 3
        MMonthName = "March"
    Case 4
        MMonthName = "April"
    Case 5
        MMonthName = "May"
    Case 6
        MMonthName = "June"
    Case 7
        MMonthName = "July"
    Case 8
        MMonthName = "Auguest"
    Case 9
        MMonthName = "September"
    Case 10
        MMonthName = "October"
    Case 11
        MMonthName = "November"
    Case 12
        MMonthName = "December"
End Select
End Function

Public Function FMonthName(Mnth As Integer)
Select Case Mnth
    Case 1
        FMonthName = "›—Ê—œÌ‰"
    Case 2
        FMonthName = "«—œÌ»Â‘ "
    Case 3
        FMonthName = "Œ—œ«œ"
    Case 4
        FMonthName = " Ì—"
    Case 5
        FMonthName = "„—œ«œ"
    Case 6
        FMonthName = "‘Â—ÌÊ—"
    Case 7
        FMonthName = "„Â—"
    Case 8
        FMonthName = "¬»«‰"
    Case 9
        FMonthName = "¬–—"
    Case 10
        FMonthName = "œÌ"
    Case 11
        FMonthName = "»Â„‰"
    Case 12
        FMonthName = "«”›‰œ"
End Select
End Function

Public Function DayOfWeek(TheDay As String)
Select Case TheDay
    Case "Saturday"
        DayOfWeek = 1
    Case "Sunday"
        DayOfWeek = 2
    Case "Monday"
        DayOfWeek = 3
    Case "Tuesday"
        DayOfWeek = 4
    Case "Wednesday"
        DayOfWeek = 5
    Case "Thursday"
        DayOfWeek = 6
    Case "Friday"
        DayOfWeek = 7
End Select
End Function

Public Function TDaysInMonth(mMon As Integer, mYear As Integer) As Integer
Select Case mMon
  Case 1 To 6
    TDaysInMonth = 31
  Case 7 To 11
    TDaysInMonth = 30
  Case 12
    If sh2mi(mYear & "/12/31") = sh2mi(mYear + 1 & "/1/1") Then
      TDaysInMonth = 30
    ElseIf sh2mi(mYear & "/12/29") = sh2mi(mYear + 1 & "/1/1") Then
        TDaysInMonth = 29
      ElseIf sh2mi(mYear & "/12/30") = sh2mi(mYear + 1 & "/1/1") Then
          TDaysInMonth = 29
    End If
End Select
End Function
Public Function TYear(TDate As Date) As Integer
TYear = Mid(mi2sh(TDate), 1, 4)
End Function
Public Function TMonth(TDate As Date) As Integer
TMonth = Mid(mi2sh(TDate), 6, 2)
End Function
Public Function TDay(TDate As Date) As Integer
TDay = Mid(mi2sh(TDate), 9, 2)
End Function

Public Function TSeq(TDay As Integer) As String
Select Case TDay
  Case 1
    TSeq = "«Ê·"
  Case 2
    TSeq = "œÊ„"
  Case 3
    TSeq = "”Ê„"
  Case 4
    TSeq = "çÂ«—„"
  Case 5
    TSeq = "Å‰Ã„"
  Case 6
    TSeq = "‘‘„"
  Case 7
    TSeq = "Â› „"
  Case 8
    TSeq = "Â‘ „"
  Case 9
    TSeq = "‰Â„"
  Case 10
    TSeq = "œÂ„"

  Case 11
    TSeq = "Ì«“œÂ„"
  Case 12
    TSeq = "œÊ«“œÂ„"
  Case 13
    TSeq = "”Ì“œÂ„"
  Case 14
    TSeq = "çÂ«—œÂ„"
  Case 15
    TSeq = "Å«‰“œÂ„"
  Case 16
    TSeq = "‘«‰“œÂ„"
  Case 17
    TSeq = "Â›œÂ„"
  Case 18
    TSeq = "ÂÃœÂ„"
  Case 19
    TSeq = "‰Ê“œÂ„"
  Case 20
    TSeq = "»Ì” „"

  Case 21
    TSeq = "»Ì  Ê Ìﬂ„"
  Case 22
    TSeq = "»Ì”  Ê œÊ„"
  Case 23
    TSeq = "»Ì”  Ê ”Ê„"
  Case 24
    TSeq = "»Ì”  Ê çÂ«—„"
  Case 25
    TSeq = "»Ì”  Ê Å‰Ã„"
  Case 26
    TSeq = "»Ì”  Ê ‘‘„"
  Case 27
    TSeq = "»Ì”  Ê Â› „"
  Case 28
    TSeq = "»Ì”  Ê Â‘ „"
  Case 29
    TSeq = "»Ì”  Ê ‰Â„"
  Case 30
    TSeq = "”Ì„"

  Case 31
    TSeq = "”Ì Ê Ìﬂ„"
End Select
End Function

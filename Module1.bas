Attribute VB_Name = "Module1"
Option Explicit
Private Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As LongPtr
Private Declare PtrSafe Function EmptyClipboard Lib "user32" () As LongPtr
Private Declare PtrSafe Function CloseClipboard Lib "user32" () As LongPtr

Private Function ClearClipboard()
  OpenClipboard (0&)

  EmptyClipboard

  CloseClipboard

End Function



Sub TimeRap()
    Dim WB, WB2     As Workbook
    Dim TimeStamp, TimeReport, Settings, Menu, Staff, Temp, DayReport As Worksheet
    Dim sourceColumn As Range, targetColumn As Range
    Dim Sht, Totals    As Worksheet
    Dim Ws_Count As Integer
    Dim Ws As Worksheet
    Dim TimeInn, TimeOut As Variant
    Dim StartTime, EndTime, Workhours As Variant
    Dim aTime  As Variant
    Dim aTemp  As Variant
    Dim aDays  As Variant
    Dim aReport As Variant
    Dim aName  As Variant
    Dim aDate  As Variant
    Dim aKalender As Variant
    Dim i, j, x, Y As Long
    Dim tall   As Integer
    Dim Kol    As Integer
    Dim Lastrow, Rad, Lastrec, Lastday As Long
    Dim Dato, CurDato As Date
    Dim Fradato, Tildato As Date
    Dim FraRad, TilRad As Long
    Dim sName  As String
    Dim NewName As String
    Dim Tid    As Double
    Dim Teller As Integer
    Dim NewHour, NewMinute, NewSecond As Variant
    Dim WaitTime As Variant
    Dim SumTekst As String
    Dim Sti As String
    Dim Filnavn As String
    Dim Hellig As Boolean


    Set WB = ThisWorkbook
    Set TimeStamp = WB.Sheets("Timestamp")
    Set Settings = WB.Sheets("Settings")
    Set TimeReport = WB.Sheets("Timereport")
    Set Temp = WB.Sheets("Temp")
    Set Menu = WB.Sheets("Menu")
    Set DayReport = WB.Sheets("Dayreport")
    
    Sti = Settings.Cells(1, 2).Value

    Application.ScreenUpdating = False
    
    CurDato = Now - 365
    Tid = Settings.Cells(10, 2).Value
    MonthFrm.Show
    
    Temp.Cells.Clear
    Fradato = Settings.Cells(7, 2).Value
    Tildato = Settings.Cells(8, 2).Value
    
    TimeStamp.Activate
    With ActiveSheet
        Lastrow = .Cells(.Rows.Count, "A").End(xlUp).Row
        Columns("A:E").Select
        ActiveWorkbook.Worksheets("TimeStamp").Sort.SortFields.Clear
        ActiveWorkbook.Worksheets("TimeStamp").Sort.SortFields.Add2 Key:=Range( _
            "C2:C" & Lastrow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
            xlSortNormal
        With ActiveWorkbook.Worksheets("TimeStamp").Sort
            .SetRange Range("A1:E1" & Lastrow)
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With
        
        aTime = Range("A1:F" & Lastrow)
        For i = 2 To Lastrow
            If aTime(i, 3) >= Fradato Then
                FraRad = i
                Exit For
            End If
        Next i
        For i = FraRad To Lastrow
            If aTime(i, 3) > Tildato Then
                TilRad = i - 1
                Exit For
            End If
        Next i
        
        aTemp = Range("A" & FraRad & ":D" & Lastrow)  ' Kan kanskje være lastrow
    End With
    
    Temp.Activate
    With ActiveSheet
        Cells(1, 1).Value = "Rec #"
        Cells(1, 2).Value = "Name"
        Cells(1, 3).Value = "Date"
        Cells(1, 4).Value = "Time In"
        Cells(1, 5).Value = "Time out"
        Columns("A:A").Select
        Selection.ColumnWidth = 7.33
        Columns("B:B").Select
        Selection.ColumnWidth = 26.33
        Selection.ColumnWidth = 19.11
        Columns("C:C").Select
        Selection.ColumnWidth = 13.67
        Selection.ColumnWidth = 11.22
        Columns("C:C").Select
        Selection.NumberFormat = "dd/mm/yyyy;@"
        Columns("D:E").Select
        Selection.ColumnWidth = 9.44
        Selection.NumberFormat = "hh:mm;@"
        
        Range("A2:D" & UBound(aTemp) + 1) = aTemp
        Lastrow = .Cells(.Rows.Count, "A").End(xlUp).Row
        
        Columns("A:E").Select
        ActiveWorkbook.Worksheets("Temp").Sort.SortFields.Clear
        ActiveWorkbook.Worksheets("Temp").Sort.SortFields.Add2 Key:=Range("B2:B" & Lastrow) _
            , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        ActiveWorkbook.Worksheets("Temp").Sort.SortFields.Add2 Key:=Range("C2:C" & Lastrow) _
            , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        With ActiveWorkbook.Worksheets("Temp").Sort
            .SetRange Range("A1:E" & Lastrow)
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With
        aReport = Range("A1:D" & Lastrow)
    End With
    
    TimeReport.Activate
    Rad = 2
    With ActiveSheet
        Range("A2:G1000").Cells.Clear
        For i = 2 To UBound(aReport)
            Cells(Rad, 1).Value = aReport(i, 1)
            Cells(Rad, 2).Value = aReport(i, 2)
            Cells(Rad, 3).Value = aReport(i, 3)
            If aReport(i, 4) > Tid Then
                Cells(Rad, 5).Value = aReport(i, 4)
            Else
                Cells(Rad, 4).Value = aReport(i, 4)
            End If
            Range("D" & i & ":E" & 1).NumberFormat = "hh:mm;@"
            If i < UBound(aReport) Then
                If aReport(i + 1, 2) = aReport(i, 2) And aReport(i + 1, 3) = aReport(i, 3) Then
                    Cells(Rad, 5).Value = aReport(i + 1, 4)
                    i = i + 1
                    tall = 1
                End If
            End If
            
            Rad = Rad + 1
        Next i
        Lastrow = .Cells(.Rows.Count, "A").End(xlUp).Row
        
        aDays = Range("A1:E" & Lastrow)
    End With
    Tildato = Settings.Cells(8, 2).Value
    DayReport.Activate
    Kol = 3
    With ActiveSheet
        DayReport.Cells.Clear
        Cells(1, 1).Value = "Weekday"
        Cells(1, 2).Value = "Date"
        Cells(1, 3).Value = "Name"
        Cells(1, 4).Value = "Time in"
        Cells(1, 5).Value = "Time out"
        Cells(1, 6).Value = "Work Hours"
        Cells(1, 7).Value = "Absent"
        Columns(Bokstav(Kol + 1) & ":" & Bokstav(Kol + 2)).Select
        Selection.ColumnWidth = 7.8
        Selection.NumberFormat = "hh:mm;@"
        Range("C:C").ColumnWidth = 15.1
        Range(Cells(1, 1), Cells(1, 8)).Interior.Color = RGB(204, 229, 255)
        aName = Range("C1:G1")
        For i = 2 To 100
            Cells(i, 1).Value = GpDay(Fradato + (i - 2))
            Cells(i, 2).Value = Fradato + (i - 2)
            If Cells(i, 2).Value >= Tildato Then Exit For
        Next i
        aDate = Range("B1:B32")
        sName = aDays(2, 2)
        
        For i = 2 To UBound(aDays)
            If Not aDays(i, 2) = sName Then
                sName = aDays(i, 2)
                Kol = Kol + 5
                Cells(1, Kol).Value = "Name"
                Cells(1, Kol + 1).Value = "Time In"
                Cells(1, Kol + 2).Value = "Time Out"
                Cells(1, Kol + 3).Value = "Work Hours"
                Cells(1, Kol + 4).Value = "Absent"
                Range(Cells(1, Kol), Cells(1, Kol + 5)).Interior.Color = RGB(204, 229, 255)
                Range(Bokstav(Kol) & ":" & Bokstav(Kol)).ColumnWidth = 15.1
                Columns(Bokstav(Kol + 1) & ":" & Bokstav(Kol + 2)).Select
                Selection.ColumnWidth = 7.8
                Selection.NumberFormat = "hh:mm;@"
            End If
            Dato = aDays(i, 3)
            Rad = FindDate(aDate, Dato)
            If Rad > 0 Then
                Cells(Rad, Kol).Value = aDays(i, 2)
                Cells(Rad, Kol + 1).Value = aDays(i, 4)
                Cells(Rad, Kol + 2).Value = aDays(i, 5)
            End If
        Next i
        
    End With
    With ActiveWindow
        .SplitColumn = 2
        .SplitRow = 0
    End With
    ActiveWindow.FreezePanes = True
    ThisWorkbook.Sheets("Menu").Activate
Application.ScreenUpdating = True
End Sub
Sub Skriv_Rapport()
    Dim WB, WB2     As Workbook
    Dim TimeStamp, TimeReport, Settings, Menu, Staff, Temp, DayReport As Worksheet
    Dim sourceColumn As Range, targetColumn As Range
    Dim Sht, Totals    As Worksheet
    Dim Ws_Count As Integer
    Dim Ws As Worksheet
    Dim TimeInn, TimeOut As Variant
    Dim StartTime, EndTime, Workhours As Variant
    Dim aTime  As Variant
    Dim aTemp  As Variant
    Dim aDays  As Variant
    Dim aReport As Variant
    Dim aName  As Variant
    Dim aDate  As Variant
    Dim aKalender As Variant
    Dim i, j, x, Y As Long
    Dim tall   As Integer
    Dim Kol    As Integer
    Dim Lastrow, Rad, Lastrec, Lastday As Long
    Dim Dato, CurDato As Date
    Dim Fradato, Tildato As Date
    Dim FraRad, TilRad As Long
    Dim sName  As String
    Dim NewName As String
    Dim Tid    As Double
    Dim Teller As Integer
    Dim NewHour, NewMinute, NewSecond As Variant
    Dim WaitTime As Variant
    Dim SumTekst As String
    Dim Sti As String
    Dim Filnavn As String
    Dim Hellig As Boolean
    Dim Nyutskrift As Boolean
    Dim Svar As String
    Dim NyBok As String
    
    Dim kildeMappe As String
    Dim destMappe As String
    Dim fil As String
    
    
    Set WB = ThisWorkbook
    Set TimeStamp = WB.Sheets("Timestamp")
    Set Settings = WB.Sheets("Settings")
    Set TimeReport = WB.Sheets("Timereport")
    Set Temp = WB.Sheets("Temp")
    Set Menu = WB.Sheets("Menu")
    Set DayReport = WB.Sheets("Dayreport")
    
    Sti = Settings.Cells(1, 2).Value
    
    Application.ScreenUpdating = False

    DayReport.Activate
    With ActiveSheet
        Lastday = DayReport.Cells(DayReport.Rows.Count, "A").End(xlUp).Row
        aKalender = Range("A1:B" & Lastday)
    End With
    
    'Svar = MsgBox("Er dette ny utskrift?", vbYesNo)

    'If Svar = vbNo Then
    Application.DisplayAlerts = False
    Workbooks.Add
    ActiveWorkbook.SaveAs Filename:=Settings.Cells(3, 2).Value & "\" & GiveName(Settings.Cells(5, 2).Value, 10)
    Application.DisplayAlerts = True
    Set WB2 = ActiveWorkbook
    'Else
    '    NyBok = GiveName(Settings.Cells(5, 2).Value, 10)
    '    Set WB2 = Workbooks(NyBok)
    '    Set Totals = WB2.Worksheets("Totals")
    '    GoTo UTSKRIFT
    'End If
    
    
    
    
    
    Set WB2 = ActiveWorkbook
    WB2.Sheets(1).Name = "Totals"
    Set Totals = WB2.Sheets("Totals")
    Totals.Cells(1, 1).Value = "Name"
    Totals.Cells(1, 2).Value = "Absent"
    Totals.Cells(1, 3).Value = "Days of lunch"
    Totals.Cells(1, 4).Value = "Lunch-Fee"
    Totals.Cells(1, 5).Value = "Hours"
    Totals.Columns("A:A").Select
    Selection.ColumnWidth = 20.1
    Totals.Columns("B:B").Select
    Selection.ColumnWidth = 8.1
    Totals.Columns("C:E").Select
    Selection.ColumnWidth = 10.1
    
    
    Totals.Range("A1:E1").Interior.Color = RGB(204, 229, 255)
    tall = 2
    For i = 3 To 100 Step 5
        For j = 2 To 35
            If Not DayReport.Cells(j, i).Value = "" Then
                NewName = DayReport.Cells(j, i).Value
                WB2.Sheets.Add after:=Sheets(Sheets.Count)
                ActiveSheet.Name = NewName
                Set Sht = WB2.Sheets(NewName)
                Totals.Cells(tall, 1).Value = NewName
                With Sht
                    Sht.Activate
                    With ActiveSheet
                        Range("A1:B" & Lastday) = aKalender
                        Columns(2).ColumnWidth = 10
                        For x = 2 To Lastday
                            Sht.Cells(x, 1).Value = DayReport.Cells(x, 1).Value
                            Sht.Cells(x, 1).NumberFormat = "hh:mm;@"
                            Sht.Cells(x, 2).Value = DayReport.Cells(x, 2).Value
                            Hellig = False
                            If DayReport.Cells(x, 1).Value = "Saturday" Or DayReport.Cells(x, 1).Value = "Sunday" Then Hellig = True
                            If DayReport.Cells(x, i + 1).Value > 0 Then
                                Sht.Cells(x, 3).Value = DayReport.Cells(x, i + 1).Value
                                Sht.Cells(x, 5).Value = DayReport.Cells(x, i + 2).Value
                            Else
                                If Hellig = False Then
                                    Sht.Cells(x, 3).Value = Settings.Cells(11, 2).Value
                                    Sht.Cells(x, 5).Value = Settings.Cells(11, 2).Value
                                End If
                            End If
                            Sht.Cells(x, 3).NumberFormat = "hh:mm;@"
                            Sht.Cells(x, 5).NumberFormat = "hh:mm;@"
                            
                        Next x
                    End With
                    
                    Cells(1, 3).Value = "Check in"
                    Cells(1, 4).Value = "Adjusted"
                    Cells(1, 5).Value = "Check out"
                    Cells(1, 6).Value = "Hours"
                    Cells(1, 7).Value = "Work Hours"
                    Cells(1, 8).Value = "Absent"
                    Cells(1, 9).Value = "T_Hours"
                    Range("A1:I1").Interior.Color = RGB(204, 229, 255)
                End With
                tall = tall + 1
                Exit For
            End If
        Next
    Next i
    WB2.Activate
    StartTime = Settings.Cells(11, 2).Value
    EndTime = Settings.Cells(12, 2).Value
    Workhours = Settings.Cells(13, 2).Value
    For i = 2 To Sheets.Count
        WB2.Sheets(i).Activate
        With ActiveSheet
            Lastrow = .Cells(.Rows.Count, "A").End(xlUp).Row
            For j = 2 To Lastrow
                If Not Cells(j, 1).Value = "Saturday" And Not Cells(j, 1).Value = "Sunday" Then
                    ' Normal arbeidstid
                    Cells(j, 6).Value = Settings.Cells(13, 2).Value
                    Cells(j, 6).NumberFormat = "[h]:mm;@"    ' "hh:mm;@"
                    If Not Cells(j, 3).Value = "" And Not Cells(j, 5) = "" Then
                        Cells(j, 7).Formula = "=Sum(E" & j & "- C" & j & ")"
                        Cells(j, 7).NumberFormat = "[h]:mm;@"    ' "hh:mm;@"
                        If Cells(j, 7).Value > Workhours Then
                            Cells(j, 8).Value = 0
                        Else
                            Cells(j, 8).Formula = "=SUM(F" & j & " - G" & j & ")"
                        End If
                        Cells(j, 8).NumberFormat = "[h]:mm;@"    ' "hh:mm;@"
                        Cells(j, 9).NumberFormat = "0.00"
                        Cells(j, 9).Formula = "=(H" & j & "-INT(H" & j & "))*24"
                        If Cells(j, 9).Value > 0 And Cells(j, 9).Value < 0.25 Then Cells(j, 9).Value = 0
                    End If
                End If
            Next j
            SumTekst = "H2+"
            For j = 3 To Lastrow
                If Not Cells(j, 3).Value = "" Then
                    SumTekst = SumTekst & "H" & j
                    If j < Lastrow Then SumTekst = SumTekst & "+"
                End If
            Next j
            SumTekst = Left(SumTekst, Len(SumTekst) - 1)
            Range("H" & Lastrow + 1).Formula = "=SUM(" & SumTekst & ")"
            Range("I" & Lastrow + 1).Formula = "=SUM(I2:I" & Lastrow & ")"
            Range("I" & Lastrow + 1).NumberFormat = "0.00"
            Range("A" & Lastrow + 1 & ":I" & Lastrow + 1).Interior.Color = RGB(204, 229, 255)
            Range("A" & Lastrow + 1).Value = "Timereport " & CStr(Settings.Cells(5, 2).Value) & " - " & ActiveSheet.Name
        End With

        Teller = 0
        For j = 2 To Lastrow
            If WB2.Sheets(i).Cells(j, 5).Value > Settings.Cells(14, 2).Value And Not WB2.Sheets(i).Cells(j, 1).Value = "Friday" Then Teller = Teller + 1
        Next j

        
        WB2.Sheets(i).Range(Bokstav(8) & Lastrow + 1).Copy
        
        With WB2.Worksheets("Totals").Range("B" & i)
            .Formula = "=" & WB2.Sheets(i).Range("H" & Lastrow + 1).Address(External:=True)
        End With
        

        With WB2.Worksheets("Totals").Range("E" & i)
            .Formula = "=" & WB2.Sheets(i).Range("I" & Lastrow + 1).Address(External:=True)
        End With
        
        
        Totals.Cells(i, 2).NumberFormat = "[h]:mm;@"
        Totals.Cells(i, 5).NumberFormat = "0.00"
        Totals.Cells(i, 3).Value = Teller
        Totals.Cells(i, 4).Formula = "=SUM(C" & i & "*" & Settings.Cells(4, 2).Value & ")"
    Next i
    Menu.Activate
    Totals.Activate
    


UTSKRIFT:
    Sti = Sti & "\" & WB2.Name
    Sti = Left(Sti, Len(Sti) - 4)
    Sti = Sti & " - "
    'MsgBox ("Skal vi fortsette?")
    WB2.Activate
    Ws_Count = WB2.Worksheets.Count
    'On Error GoTo Feil
    For i = 1 To Ws_Count
        WB2.Worksheets(i).Activate
        If Not WB2.Worksheets(i).Name = "Totals" Then
            Sti = "C:\Temp"
            Sti = Sti & "\" & WB2.Name
            Sti = Left(Sti, Len(Sti) - 4)
            Sti = Sti & " - "
            Filnavn = Sti & WB2.Worksheets(i).Name & ".PDF"
            Debug.Print "Filnavn: " & Filnavn
            
            'Filnavn = Sti & WB2.Worksheets(i).Name & ".PDF"
        
            With WB2.Worksheets(i)
                With WB2.Worksheets(i).PageSetup
                    .PrintArea = "$A$1:$H$36"
                    .Zoom = False
                    .FitToPagesWide = 1
                    .FitToPagesTall = 1
                End With
                
                
                .Activate
        
                DoEvents
        
                ActiveSheet.ExportAsFixedFormat _
                    Type:=xlTypePDF, _
                    Filename:=Filnavn, _
                    Quality:=xlQualityStandard, _
                    IncludeDocProperties:=True, _
                    IgnorePrintAreas:=True, _
                    OpenAfterPublish:=False


            End With

        End If
    Next i
    
    
    ' Kilde
    kildeMappe = "C:\Temp\"
    
    ' Destinasjon fra Settings!B1
    destMappe = Settings.Cells(1, 2).Value
    
    ' Sørg for at det er \ på slutten
    If Right(destMappe, 1) <> "\" Then
        destMappe = destMappe & "\"
    End If
    
    ' Sjekk at destinasjonsmappe finnes
    If Dir(destMappe, vbDirectory) = "" Then
        MsgBox "Destinasjonsmappe finnes ikke:" & vbCrLf & destMappe, vbExclamation
        Exit Sub
    End If
    
    ' Finn første fil
    fil = Dir(kildeMappe & "*.*")
    
    Do While fil <> ""
        
        ' Flytt fil
        Name kildeMappe & fil As destMappe & fil
        
        fil = Dir
    Loop
    
    MsgBox "Alle filer er flyttet.", vbInformation
    Exit Sub

Feil:
    MsgBox "Feil: " & Err.Number & vbCrLf & Err.Description, vbCritical


    
    Totals.Activate
    With ActiveSheet
        Lastrow = Totals.Cells(Totals.Rows.Count, "A").End(xlUp).Row + 1
        Totals.Cells(Lastrow, 1).Value = "Ebrahima Badjie"
        Totals.Cells(Lastrow, 4).Value = Totals.Cells(Lastrow - 1, 4).Value
        Application.ScreenUpdating = True
    End With
    'Feil:
    '    MsgBox "Feilnummer: " & Err.Number & vbCrLf & _
    '           "Feilbeskrivelse: " & Err.Description, vbCritical

End Sub



Function GiveName(mNo As Integer, Nummer As Integer) As String
    Dim mName  As String
    If mNo = 1 Then mName = "January"
    If mNo = 2 Then mName = "February"
    If mNo = 3 Then mName = "March"
    If mNo = 4 Then mName = "April"
    If mNo = 5 Then mName = "May"
    If mNo = 6 Then mName = "June"
    If mNo = 7 Then mName = "July"
    If mNo = 8 Then mName = "August"
    If mNo = 9 Then mName = "September"
    If mNo = 10 Then mName = "October"
    If mNo = 11 Then mName = "November"
    If mNo = 12 Then mName = "December"
    If Nummer <> 2 Then
        GiveName = "Time Report " & mName & " " & Str(ThisWorkbook.Sheets("Settings").Cells(5, 2).Value) & ".xlsx"
    Else
        GiveName = "Time Report " & mName & " " & Str(ThisWorkbook.Sheets("Settings").Cells(5, 2).Value)
    End If
    
End Function

Function Bokstav(Kolonne As Variant) As String
    Dim ColumnNumber As Long
    Dim ColumnLetter As String
    ColumnNumber = Kolonne
    Bokstav = Split(Cells(1, ColumnNumber).Address, "$")(1)
    
End Function



Function GpDay(gDay As Date) As String
    Dim nDay   As Integer

    nDay = Weekday(gDay)
    
    If nDay = 2 Then GpDay = "Monday"
    If nDay = 3 Then GpDay = "Tuesday"
    If nDay = 4 Then GpDay = "Wednesday"
    If nDay = 5 Then GpDay = "Thursday"
    If nDay = 6 Then GpDay = "Friday"
    If nDay = 7 Then GpDay = "Saturday"
    If nDay = 1 Then GpDay = "Sunday"
    
    
End Function
Public Function FindDate(aRR, val As Variant) As Long
    Dim r      As Long
    For r = LBound(aRR) To UBound(aRR)
        If aRR(r, 1) = val Then
            FindDate = r
            Exit Function
        End If
    Next r
    
End Function



Sub ReadTimeFile()

    Dim WB     As Workbook
    Dim Ws     As Worksheet
    Dim Settings, Menu As Worksheet
    Dim Temp   As Worksheet
    Dim TimeStamp As Worksheet
    Dim line   As String, Filename As String, i As Integer, valueArr() As String
    Dim lPos, Ant   As Integer
    Dim År, Måned, Dag As Integer
    Dim aTemp  As Variant
    Dim aStaff As Variant
    Dim aTemp2 As Variant
    Dim aImport As Variant
    Dim sDato As String
    Dim dDato As Date
    Dim Lastrow As Long
    Dim j, x   As Integer
    Dim Sistepos As Integer
    Dim Årstall As Variant
    Dim Tekst  As String
    Dim sSti   As String
    Dim Eyvind As String
    Dim CurDato As Date
    Dim Tid As Double
    Dim Fradato, Tildato, StampDate As Date
    Dim Rad As Integer

    Set WB = ThisWorkbook
    Set Settings = WB.Sheets("Settings")
    Set Temp = WB.Sheets("Temp")
    Set Menu = WB.Sheets("Menu")
    Set TimeStamp = WB.Sheets("Timestamp")
    
    
    CurDato = Now - 365
    Tid = Settings.Cells(10, 2).Value
    MonthFrm.Show
    
    Fradato = Settings.Cells(7, 2).Value
    Tildato = Settings.Cells(8, 2).Value
    
    Application.ScreenUpdating = False
    
    sSti = Settings.Cells(2, 2).Value
    Filename = Filvelger(sSti)
    Settings.Activate
    With ActiveSheet
        Lastrow = .Cells(.Rows.Count, "E").End(xlUp).Row
        aStaff = Range("E1:F" & Lastrow)
    End With
    
    ThisWorkbook.Sheets("Temp").Cells.Clear
    
    i = 1
    Open Filename For Input As #2
    While Not EOF(2)
        Line Input #2, line
        i = i + 1
        ThisWorkbook.Sheets("Temp").Cells(i, 1).Value = line
    Wend
    'Close File
    Close #2
    ThisWorkbook.Sheets("Temp").Activate
    With ActiveSheet
        Lastrow = .Cells(.Rows.Count, "A").End(xlUp).Row
        
        aImport = Range("A1:B" & Lastrow)
        ThisWorkbook.Sheets("Temp").Cells.Clear
        Rad = 1
        For i = 3 To Lastrow
            'Tekst = Cells(i, 1).Value
            Tekst = aImport(i, 1)
            Ant = Len(Tekst)
            sDato = Mid(Tekst, Ant - 19, 11)
            År = CInt(Left(sDato, 5))
            Måned = CInt(Mid(sDato, 7, 2))
            Dag = CInt(Right(sDato, 2))
            dDato = DateSerial(År, Måned, Dag)
            
            Årstall = CInt(Mid(Tekst, Ant - 19, 5))
            
            'Cells(Rad, 2).Value = Ant  ' Antall pos
            aImport(i, 2) = Ant  ' Antall pos
            'If i = 102 Then Stop
            If dDato >= Fradato And dDato <= Tildato Then
            Rad = Rad + 1
            lPos = 0
            If i >= 102 Then lPos = 1
            If i >= 1002 Then lPos = 2
            If i >= 1002 Then lPos = 3
            For j = Ant To 1 Step -1
                If Mid(Tekst, j, 1) = ":" Then
                    Cells(Rad, 3).Value = j  ' Pos til kolon
                    Cells(Rad, 4).Value = Mid(Tekst, j - 6, 9)  ' Tid
                    Cells(Rad, 5).Value = Mid(Tekst, j - 17, 5) ' År
                    Cells(Rad, 6).Value = Mid(Tekst, j - 11, 2) ' Måned
                    Cells(Rad, 7).Value = Mid(Tekst, j - 8, 2)  ' Dag
                    
                    Cells(Rad, 8).Value = val(Mid(Tekst, 6 + lPos, 8))    ' Ansattnr
                    Cells(Rad, 9).Value = Left(Tekst, 5)        ' Rec nummer
                    Exit For
                End If
            Next j
            End If
        Next i
        Lastrow = .Cells(.Rows.Count, "C").End(xlUp).Row
        aTemp = Range("A1:I" & Lastrow)
        'ThisWorkbook.Sheets("Temp").Cells.Clear
    End With
    TimeStamp.Activate
    With ActiveSheet
        Range("A2:E50000").Clear
        Rad = 2
        For i = 2 To UBound(aTemp)
            'If i = 3000 Then Stop
            StampDate = DateSerial(CInt(aTemp(i, 5)), CInt(aTemp(i, 6)), CInt(aTemp(i, 7)))
            If StampDate >= Fradato And StampDate <= Tildato Then
                Cells(Rad, 1).Value = val(aTemp(i, 9)) ' Rec number
                Cells(Rad, 2).Value = aStaff(aTemp(i, 8) + 1, 2) ' Navn
                If Not aTemp(i, 4) = "" Then
                    Cells(Rad, 3).Value = StampDate
                    Cells(Rad, 4).Value = TimeSerial(CInt(Mid(aTemp(i, 4), 2, 2)), _
                        CLng(Mid(aTemp(i, 4), 5, 2)), _
                        CLng(Right(aTemp(i, 4), 2)))
                End If
                Rad = Rad + 1
            End If
            
        Next i
    End With
    Menu.Activate
    Application.ScreenUpdating = True
End Sub





Public Function Filvelger(Oppstart As String) As String

    'Declare a variable as a FileDialog object
    Dim fd     As FileDialog

    'Create a FileDialog object as a File Picker dialog box.
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    
    Dim vrtSelectedItem As Variant
    
    With fd
        
        .InitialFileName = Oppstart
        
        If .Show = -1 Then
            
            For Each vrtSelectedItem In .SelectedItems
                Filvelger = vrtSelectedItem
                Exit Function
            Next vrtSelectedItem
            'If the user presses Cancel...
        Else
        End If
    End With
    
    'Set the object variable to Nothing.
    Set fd = Nothing
    
End Function

Function FileExists(fname) As Boolean
    ' Sjekker om en fil finnes. Send sti og filnavn
    If Dir(fname) <> "" Then _
    FileExists = True _
Else: FileExists = False
End Function

Public Function FindStaff(aRR, val As Variant) As Long
    Dim r      As Long
    For r = LBound(aRR) To UBound(aRR)
        If aRR(r, 1) = val Then
            FindStaff = r
            Exit Function
        End If
    Next r
    
End Function

Sub Test()
    Dim ma     As Date
    Dim sMa    As String
    Dim i      As Integer
    ma = Now
    For i = 1 To 12
        sMa = LagStorBokstav(MonthName(i))
        ThisWorkbook.Sheets("Settings").Cells(i + 1, 3).Value = sMa
    Next i
    
End Sub

Function Maaned(Inn As Integer)
    Maaned = Month(Inn)
End Function

Function MonthNumber(myMonthName As String)
    Dim sYear  As String
    sYear = " " & Str(ThisWorkbook.Sheets("Settings").Cells(6, 2).Value)
    MonthNumber = Month(DateValue("1 " & myMonthName & sYear))
    MonthNumber = Format(MonthNumber, "00")
End Function

Public Function LagStorBokstav(sWord As String) As String
    ' Endrer første bokstav til uppercase
    LagStorBokstav = StrConv(sWord, vbProperCase)
End Function


Public Function TallMedNull(lTall As Long, iAntall As Integer) As String
    Dim i      As Integer
    Dim sNull  As String
    sNull = ""
    If iAntall > 0 Then
        For i = 1 To iAntall
            sNull = sNull & "0"
        Next i
    End If
    TallMedNull = Format(lTall, sNull)
End Function

Public Function VelgFarge(Kode As String) As Long
    ' Rød = 255,204,255
    ' Lysere grønn = 229,255,204
    ' Mørkere grønn = 204, 255, 204
    Dim fKode  As String
    fKode = Right(Kode, 1)
    
    If fKode = "0" Then
        VelgFarge = RGB(224, 224, 224)            ' Grå
        'Exit Function
    End If
    If fKode = "1" Then
        VelgFarge = RGB(255, 255, 204)            ' Gul
        'Exit Function
    End If
    
    If fKode = "2" Then
        VelgFarge = RGB(229, 255, 204)            ' Grønn
        'Exit Function
    End If
    fKode = Right(Kode, 3)
    If fKode = "ett" Then
        VelgFarge = RGB(204, 229, 255)            ' Blå
        'Exit Function
    End If
    fKode = "Butikker"
    If fKode = "Butikker" Then
        VelgFarge = RGB(204, 229, 255)            ' Blå
        'Exit Function
    End If
    
    If VelgFarge = 0 Then VelgFarge = RGB(255, 255, 255)
End Function


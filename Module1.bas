Attribute VB_Name = "Module1"
Attribute VB_Name = "Module1"
Option Explicit

Private Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As LongPtr
Private Declare PtrSafe Function EmptyClipboard Lib "user32" () As LongPtr
Private Declare PtrSafe Function CloseClipboard Lib "user32" () As LongPtr

' ============================================================
'  Hjelpefunksjoner
' ============================================================

Private Sub ClearClipboard()
    OpenClipboard 0&
    EmptyClipboard
    CloseClipboard
End Sub

' ------------------------------------------------------------
'  Felles oppsett: henter ark-referanser og datoperiode
' ------------------------------------------------------------
Private Sub InitWorksheets( _
        ByRef WB As Workbook, _
        ByRef TimeStamp As Worksheet, _
        ByRef Settings As Worksheet, _
        ByRef TimeReport As Worksheet, _
        ByRef Temp As Worksheet, _
        ByRef Menu As Worksheet, _
        ByRef DayReport As Worksheet)

    Set WB = ThisWorkbook
    Set TimeStamp = WB.Sheets("Timestamp")
    Set Settings = WB.Sheets("Settings")
    Set TimeReport = WB.Sheets("Timereport")
    Set Temp = WB.Sheets("Temp")
    Set Menu = WB.Sheets("Menu")
    Set DayReport = WB.Sheets("Dayreport")
End Sub

' ============================================================
'  TimeRap  –  sorterer og bygger dagrapport
' ============================================================
Sub TimeRap()
    ' --- Arbeidsbok og ark ---
    Dim WB         As Workbook
    Dim WB2        As Workbook
    Dim TimeStamp  As Worksheet
    Dim TimeReport As Worksheet
    Dim Settings   As Worksheet
    Dim Menu       As Worksheet
    Dim Temp       As Worksheet
    Dim DayReport  As Worksheet

    ' --- Område og arrays ---
    Dim aTime      As Variant
    Dim aTemp      As Variant
    Dim aDays      As Variant
    Dim aReport    As Variant
    Dim aDate      As Variant

    ' --- Løkkevariabler ---
    Dim i          As Long
    Dim tall       As Integer
    Dim Kol        As Integer
    Dim Lastrow    As Long
    Dim Rad        As Long
    Dim FraRad     As Long

    ' --- Dato og tid ---
    Dim Dato       As Date
    Dim Fradato    As Date
    Dim Tildato    As Date
    Dim Tid        As Double

    ' --- Strenger ---
    Dim sName      As String

    InitWorksheets WB, TimeStamp, Settings, TimeReport, Temp, Menu, DayReport

    Application.ScreenUpdating = False

    Tid = Settings.Cells(10, 2).Value
    MonthFrm.Show

    Temp.Cells.Clear
    Fradato = Settings.Cells(7, 2).Value
    Tildato = Settings.Cells(8, 2).Value

    ' --- 1. Les og sorter TimeStamp-ark ---
    With TimeStamp
        Lastrow = .Cells(.Rows.Count, "A").End(xlUp).Row

        With .Sort
            .SortFields.Clear
            .SortFields.Add2 Key:=TimeStamp.Range("C2:C" & Lastrow), _
                             SortOn:=xlSortOnValues, Order:=xlAscending, _
                             DataOption:=xlSortNormal
            .SetRange TimeStamp.Range("A1:E" & Lastrow)   ' FIX: var "A1:E1" & Lastrow
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With

        aTime = .Range("A1:F" & Lastrow)

        ' Finn første rad >= Fradato
        For i = 2 To Lastrow
            If aTime(i, 3) >= Fradato Then
                FraRad = i
                Exit For
            End If
        Next i

        ' Les ut ønsket periode (til og med Lastrow)
        aTemp = .Range("A" & FraRad & ":D" & Lastrow)
    End With

    ' --- 2. Bygg Temp-ark ---
    With Temp
        .Cells(1, 1).Value = "Rec #"
        .Cells(1, 2).Value = "Name"
        .Cells(1, 3).Value = "Date"
        .Cells(1, 4).Value = "Time In"
        .Cells(1, 5).Value = "Time out"

        .Columns("A").ColumnWidth = 7.33
        .Columns("B").ColumnWidth = 19.11
        .Columns("C").ColumnWidth = 11.22
        .Columns("C").NumberFormat = "dd/mm/yyyy;@"
        .Columns("D:E").ColumnWidth = 9.44
        .Columns("D:E").NumberFormat = "hh:mm;@"

        .Range("A2:D" & UBound(aTemp) + 1) = aTemp
        Lastrow = .Cells(.Rows.Count, "A").End(xlUp).Row

        With .Sort
            .SortFields.Clear
            .SortFields.Add2 Key:=Temp.Range("B2:B" & Lastrow), _
                             SortOn:=xlSortOnValues, Order:=xlAscending, _
                             DataOption:=xlSortNormal
            .SortFields.Add2 Key:=Temp.Range("C2:C" & Lastrow), _
                             SortOn:=xlSortOnValues, Order:=xlAscending, _
                             DataOption:=xlSortNormal
            .SetRange Temp.Range("A1:E" & Lastrow)
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With

        aReport = .Range("A1:D" & Lastrow)
    End With

    ' --- 3. Bygg TimeReport-ark ---
    With TimeReport
        .Range("A2:G1000").Cells.Clear
        Rad = 2
        For i = 2 To UBound(aReport)
            .Cells(Rad, 1).Value = aReport(i, 1)
            .Cells(Rad, 2).Value = aReport(i, 2)
            .Cells(Rad, 3).Value = aReport(i, 3)

            If aReport(i, 4) > Tid Then
                .Cells(Rad, 5).Value = aReport(i, 4)
            Else
                .Cells(Rad, 4).Value = aReport(i, 4)
            End If

            .Range("D" & Rad & ":E" & Rad).NumberFormat = "hh:mm;@"   ' FIX: var ":E" & 1

            ' Kombiner inn/ut for samme person og dag
            If i < UBound(aReport) Then
                If aReport(i + 1, 2) = aReport(i, 2) And _
                   aReport(i + 1, 3) = aReport(i, 3) Then
                    .Cells(Rad, 5).Value = aReport(i + 1, 4)
                    i = i + 1
                    tall = 1
                End If
            End If

            Rad = Rad + 1
        Next i

        Lastrow = .Cells(.Rows.Count, "A").End(xlUp).Row
        aDays = .Range("A1:E" & Lastrow)
    End With

    ' --- 4. Bygg DayReport-ark ---
    Tildato = Settings.Cells(8, 2).Value
    Kol = 3

    With DayReport
        .Cells.Clear
        .Cells(1, 1).Value = "Weekday"
        .Cells(1, 2).Value = "Date"
        .Cells(1, 3).Value = "Name"
        .Cells(1, 4).Value = "Time in"
        .Cells(1, 5).Value = "Time out"
        .Cells(1, 6).Value = "Work Hours"
        .Cells(1, 7).Value = "Absent"

        .Columns(Bokstav(Kol + 1) & ":" & Bokstav(Kol + 2)).ColumnWidth = 7.8
        .Columns(Bokstav(Kol + 1) & ":" & Bokstav(Kol + 2)).NumberFormat = "hh:mm;@"
        .Columns("C").ColumnWidth = 15.1
        .Range(.Cells(1, 1), .Cells(1, 8)).Interior.Color = RGB(204, 229, 255)

        ' Fyll inn datokolonne
        For i = 2 To 100
            .Cells(i, 1).Value = GpDay(Fradato + (i - 2))
            .Cells(i, 2).Value = Fradato + (i - 2)
            If .Cells(i, 2).Value >= Tildato Then Exit For
        Next i

        aDate = .Range("B1:B32")
        sName = aDays(2, 2)

        ' Fyll inn per ansatt og dato
        For i = 2 To UBound(aDays)
            If Not aDays(i, 2) = sName Then
                sName = aDays(i, 2)
                Kol = Kol + 5
                .Cells(1, Kol).Value = "Name"
                .Cells(1, Kol + 1).Value = "Time In"
                .Cells(1, Kol + 2).Value = "Time Out"
                .Cells(1, Kol + 3).Value = "Work Hours"
                .Cells(1, Kol + 4).Value = "Absent"
                .Range(.Cells(1, Kol), .Cells(1, Kol + 5)).Interior.Color = RGB(204, 229, 255)
                .Columns(Bokstav(Kol)).ColumnWidth = 15.1
                .Columns(Bokstav(Kol + 1) & ":" & Bokstav(Kol + 2)).ColumnWidth = 7.8
                .Columns(Bokstav(Kol + 1) & ":" & Bokstav(Kol + 2)).NumberFormat = "hh:mm;@"
            End If

            Dato = aDays(i, 3)
            Rad = FindDate(aDate, Dato)
            If Rad > 0 Then
                .Cells(Rad, Kol).Value = aDays(i, 2)
                .Cells(Rad, Kol + 1).Value = aDays(i, 4)
                .Cells(Rad, Kol + 2).Value = aDays(i, 5)
            End If
        Next i
    End With

    ' Frys de to første kolonnene
    DayReport.Activate
    With ActiveWindow
        .SplitColumn = 2
        .SplitRow = 0
    End With
    ActiveWindow.FreezePanes = True

    Menu.Activate
    Application.ScreenUpdating = True
End Sub


' ============================================================
'  Skriv_Rapport  –  lager PDF-rapport for hver ansatt
' ============================================================
Sub Skriv_Rapport()
    ' --- Arbeidsbok og ark ---
    Dim WB         As Workbook
    Dim WB2        As Workbook
    Dim TimeStamp  As Worksheet
    Dim TimeReport As Worksheet
    Dim Settings   As Worksheet
    Dim Menu       As Worksheet
    Dim Temp       As Worksheet
    Dim DayReport  As Worksheet
    Dim Sht        As Worksheet
    Dim Totals     As Worksheet

    ' --- Arrays ---
    Dim aKalender  As Variant
    Dim aDays      As Variant

    ' --- Løkkevariabler ---
    Dim i          As Long
    Dim j          As Long
    Dim x          As Long
    Dim Kol        As Integer
    Dim Teller     As Integer
    Dim tall       As Integer

    ' --- Størrelser ---
    Dim Lastrow    As Long
    Dim Lastday    As Long
    Dim Ws_Count   As Integer

    ' --- Tid ---
    Dim StartTime  As Variant
    Dim EndTime    As Variant
    Dim Workhours  As Variant

    ' --- Strenger ---
    Dim NewName    As String
    Dim SumTekst   As String
    Dim Sti        As String
    Dim Filnavn    As String

    ' --- Flagg ---
    Dim Hellig     As Boolean

    InitWorksheets WB, TimeStamp, Settings, TimeReport, Temp, Menu, DayReport

    Application.ScreenUpdating = False

    ' Hent kalenderdata fra DayReport
    With DayReport
        Lastday = .Cells(.Rows.Count, "A").End(xlUp).Row
        aKalender = .Range("A1:B" & Lastday)
    End With

    ' Hent dagdata fra DayReport (bygd av TimeRap)
    With DayReport
        aKalender = .Range("A1:B" & Lastday)
    End With

    ' --- Opprett ny arbeidsbok for rapporten ---
    Application.DisplayAlerts = False
    Workbooks.Add
    ActiveWorkbook.SaveAs Filename:= _
        Settings.Cells(3, 2).Value & "\" & GiveName(Settings.Cells(5, 2).Value, 10)
    Application.DisplayAlerts = True
    Set WB2 = ActiveWorkbook

    WB2.Sheets(1).Name = "Totals"
    Set Totals = WB2.Sheets("Totals")

    ' Sett opp Totals-ark
    With Totals
        .Cells(1, 1).Value = "Name"
        .Cells(1, 2).Value = "Absent"
        .Cells(1, 3).Value = "Days of lunch"
        .Cells(1, 4).Value = "Lunch-Fee"
        .Cells(1, 5).Value = "Hours"
        .Columns("A").ColumnWidth = 20.1
        .Columns("B").ColumnWidth = 8.1
        .Columns("C:E").ColumnWidth = 10.1
        .Range("A1:E1").Interior.Color = RGB(204, 229, 255)
    End With

    ' --- Opprett ett ark per ansatt ---
    tall = 2
    For i = 3 To 100 Step 5
        For j = 2 To 35
            If Not DayReport.Cells(j, i).Value = "" Then
                NewName = DayReport.Cells(j, i).Value
                WB2.Sheets.Add After:=WB2.Sheets(WB2.Sheets.Count)
                ActiveSheet.Name = NewName
                Set Sht = WB2.Sheets(NewName)
                Totals.Cells(tall, 1).Value = NewName

                With Sht
                    .Range("A1:B" & Lastday) = aKalender
                    .Columns(2).ColumnWidth = 10

                    For x = 2 To Lastday
                        .Cells(x, 1).Value = DayReport.Cells(x, 1).Value
                        .Cells(x, 1).NumberFormat = "hh:mm;@"
                        .Cells(x, 2).Value = DayReport.Cells(x, 2).Value

                        Hellig = (DayReport.Cells(x, 1).Value = "Saturday" Or _
                                  DayReport.Cells(x, 1).Value = "Sunday")

                        If DayReport.Cells(x, i + 1).Value > 0 Then
                            .Cells(x, 3).Value = DayReport.Cells(x, i + 1).Value
                            .Cells(x, 5).Value = DayReport.Cells(x, i + 2).Value
                        ElseIf Not Hellig Then
                            .Cells(x, 3).Value = Settings.Cells(11, 2).Value
                            .Cells(x, 5).Value = Settings.Cells(11, 2).Value
                        End If

                        .Cells(x, 3).NumberFormat = "hh:mm;@"
                        .Cells(x, 5).NumberFormat = "hh:mm;@"
                    Next x

                    .Cells(1, 3).Value = "Check in"
                    .Cells(1, 4).Value = "Adjusted"
                    .Cells(1, 5).Value = "Check out"
                    .Cells(1, 6).Value = "Hours"
                    .Cells(1, 7).Value = "Work Hours"
                    .Cells(1, 8).Value = "Absent"
                    .Cells(1, 9).Value = "T_Hours"
                    .Range("A1:I1").Interior.Color = RGB(204, 229, 255)
                End With

                tall = tall + 1
                Exit For
            End If
        Next j
    Next i

    ' --- Beregn timer per ansatt ---
    StartTime = Settings.Cells(11, 2).Value
    EndTime = Settings.Cells(12, 2).Value
    Workhours = Settings.Cells(13, 2).Value

    Ws_Count = WB2.Worksheets.Count
    For i = 2 To Ws_Count
        With WB2.Sheets(i)
            Lastrow = .Cells(.Rows.Count, "A").End(xlUp).Row

            For j = 2 To Lastrow
                If Not .Cells(j, 1).Value = "Saturday" And _
                   Not .Cells(j, 1).Value = "Sunday" Then
                    .Cells(j, 6).Value = Workhours
                    .Cells(j, 6).NumberFormat = "[h]:mm;@"

                    If Not .Cells(j, 3).Value = "" And Not .Cells(j, 5).Value = "" Then
                        .Cells(j, 7).Formula = "=SUM(E" & j & "-C" & j & ")"
                        .Cells(j, 7).NumberFormat = "[h]:mm;@"

                        If .Cells(j, 7).Value > Workhours Then
                            .Cells(j, 8).Value = 0
                        Else
                            .Cells(j, 8).Formula = "=SUM(F" & j & "-G" & j & ")"
                        End If
                        .Cells(j, 8).NumberFormat = "[h]:mm;@"
                        .Cells(j, 9).NumberFormat = "0.00"
                        .Cells(j, 9).Formula = "=(H" & j & "-INT(H" & j & "))*24"
                        If .Cells(j, 9).Value > 0 And .Cells(j, 9).Value < 0.25 Then
                            .Cells(j, 9).Value = 0
                        End If
                    End If
                End If
            Next j

            ' Bygg SUM-formel for fraværskolonnen
            SumTekst = "H2"
            For j = 3 To Lastrow
                If Not .Cells(j, 3).Value = "" Then
                    SumTekst = SumTekst & "+H" & j
                End If
            Next j
            .Range("H" & Lastrow + 1).Formula = "=SUM(" & SumTekst & ")"
            .Range("I" & Lastrow + 1).Formula = "=SUM(I2:I" & Lastrow & ")"
            .Range("I" & Lastrow + 1).NumberFormat = "0.00"
            .Range("A" & Lastrow + 1 & ":I" & Lastrow + 1).Interior.Color = RGB(204, 229, 255)
            .Range("A" & Lastrow + 1).Value = "Timereport " & _
                CStr(Settings.Cells(5, 2).Value) & " - " & .Name

            ' Tell lunsjedager
            Teller = 0
            For j = 2 To Lastrow
                If .Cells(j, 5).Value > Settings.Cells(14, 2).Value And _
                   Not .Cells(j, 1).Value = "Friday" Then
                    Teller = Teller + 1
                End If
            Next j

            ' Fyll Totals-referanser
            With WB2.Worksheets("Totals")
                .Range("B" & i).Formula = "=" & WB2.Sheets(i).Range("H" & Lastrow + 1).Address(External:=True)
                .Range("E" & i).Formula = "=" & WB2.Sheets(i).Range("I" & Lastrow + 1).Address(External:=True)
                .Cells(i, 2).NumberFormat = "[h]:mm;@"
                .Cells(i, 5).NumberFormat = "0.00"
                .Cells(i, 3).Value = Teller
                .Cells(i, 4).Formula = "=SUM(C" & i & "*" & Settings.Cells(4, 2).Value & ")"
            End With
        End With
    Next i

    ' --- Eksporter PDF per ark ---
    Sti = Settings.Cells(1, 2).Value   ' FIX: bruker Settings-verdi, ikke hardkodet "C:\Temp"
    If Right(Sti, 1) <> "\" Then Sti = Sti & "\"

    For i = 1 To Ws_Count
        If Not WB2.Worksheets(i).Name = "Totals" Then
            Filnavn = Sti & Left(WB2.Name, Len(WB2.Name) - 4) & " - " & _
                      WB2.Worksheets(i).Name & ".PDF"

            With WB2.Worksheets(i)
                With .PageSetup
                    .PrintArea = "$A$1:$H$36"
                    .Zoom = False
                    .FitToPagesWide = 1
                    .FitToPagesTall = 1
                End With

                .ExportAsFixedFormat _
                    Type:=xlTypePDF, _
                    Filename:=Filnavn, _
                    Quality:=xlQualityStandard, _
                    IncludeDocProperties:=True, _
                    IgnorePrintAreas:=True, _
                    OpenAfterPublish:=False
            End With
        End If
    Next i

    ' --- Flytt PDF-filer til destinasjonsmappe ---
    Dim kildeMappe As String
    Dim destMappe  As String
    Dim fil        As String

    kildeMappe = Sti
    destMappe = Settings.Cells(1, 2).Value
    If Right(destMappe, 1) <> "\" Then destMappe = destMappe & "\"

    If Dir(destMappe, vbDirectory) = "" Then
        MsgBox "Destinasjonsmappe finnes ikke:" & vbCrLf & destMappe, vbExclamation
        GoTo Rydding
    End If

    fil = Dir(kildeMappe & "*.PDF")
    Do While fil <> ""
        Name kildeMappe & fil As destMappe & fil
        fil = Dir
    Loop
    MsgBox "Alle filer er flyttet.", vbInformation

Rydding:
    Totals.Activate
    Menu.Activate
    Application.ScreenUpdating = True
End Sub


' ============================================================
'  ReadTimeFile  –  importerer tidsstempel-fil
' ============================================================
Sub ReadTimeFile()
    Dim WB         As Workbook
    Dim Settings   As Worksheet
    Dim Menu       As Worksheet
    Dim Temp       As Worksheet
    Dim TimeStamp  As Worksheet

    Dim line       As String
    Dim Filename   As String
    Dim i          As Long
    Dim j          As Long
    Dim lPos       As Integer
    Dim Ant        As Integer
    Dim År         As Integer
    Dim Måned      As Integer
    Dim Dag        As Integer
    Dim aTemp      As Variant
    Dim aStaff     As Variant
    Dim aImport    As Variant
    Dim sDato      As String
    Dim dDato      As Date
    Dim Lastrow    As Long
    Dim Rad        As Integer
    Dim Tekst      As String
    Dim sSti       As String
    Dim Årstall    As Variant
    Dim Fradato    As Date
    Dim Tildato    As Date
    Dim StampDate  As Date

    Set WB = ThisWorkbook
    Set Settings = WB.Sheets("Settings")
    Set Temp = WB.Sheets("Temp")
    Set Menu = WB.Sheets("Menu")
    Set TimeStamp = WB.Sheets("Timestamp")

    Dim Tid As Double
    Tid = Settings.Cells(10, 2).Value
    MonthFrm.Show

    Fradato = Settings.Cells(7, 2).Value
    Tildato = Settings.Cells(8, 2).Value

    Application.ScreenUpdating = False

    sSti = Settings.Cells(2, 2).Value
    Filename = Filvelger(sSti)

    ' Les ansatte fra Settings
    Settings.Activate
    With Settings
        Lastrow = .Cells(.Rows.Count, "E").End(xlUp).Row
        aStaff = .Range("E1:F" & Lastrow)
    End With

    ' Les tekstfil linje for linje inn i Temp
    WB.Sheets("Temp").Cells.Clear
    i = 1
    Open Filename For Input As #2
    While Not EOF(2)
        Line Input #2, line
        i = i + 1
        WB.Sheets("Temp").Cells(i, 1).Value = line
    Wend
    Close #2

    ' Parse Temp og trekk ut feltene
    With WB.Sheets("Temp")
        Lastrow = .Cells(.Rows.Count, "A").End(xlUp).Row
        aImport = .Range("A1:B" & Lastrow)
        .Cells.Clear

        Rad = 1
        For i = 3 To UBound(aImport)
            Tekst = aImport(i, 1)
            Ant = Len(Tekst)
            sDato = Mid(Tekst, Ant - 19, 11)
            År = CInt(Left(sDato, 5))
            Måned = CInt(Mid(sDato, 7, 2))
            Dag = CInt(Right(sDato, 2))
            dDato = DateSerial(År, Måned, Dag)

            If dDato >= Fradato And dDato <= Tildato Then
                Rad = Rad + 1

                ' FIX: lPos-logikk var feil (to identiske grenser for lPos=3)
                lPos = 0
                If i >= 102 Then lPos = 1
                If i >= 1002 Then lPos = 2
                If i >= 10002 Then lPos = 3

                For j = Ant To 1 Step -1
                    If Mid(Tekst, j, 1) = ":" Then
                        .Cells(Rad, 3).Value = j
                        .Cells(Rad, 4).Value = Mid(Tekst, j - 6, 9)
                        .Cells(Rad, 5).Value = Mid(Tekst, j - 17, 5)
                        .Cells(Rad, 6).Value = Mid(Tekst, j - 11, 2)
                        .Cells(Rad, 7).Value = Mid(Tekst, j - 8, 2)
                        .Cells(Rad, 8).Value = val(Mid(Tekst, 6 + lPos, 8))
                        .Cells(Rad, 9).Value = Left(Tekst, 5)
                        Exit For
                    End If
                Next j
            End If
        Next i

        Lastrow = .Cells(.Rows.Count, "C").End(xlUp).Row
        aTemp = .Range("A1:I" & Lastrow)
    End With

    ' Skriv til TimeStamp-ark
    With TimeStamp
        .Range("A2:E50000").Clear
        Rad = 2
        For i = 2 To UBound(aTemp)
            StampDate = DateSerial(CInt(aTemp(i, 5)), CInt(aTemp(i, 6)), CInt(aTemp(i, 7)))
            If StampDate >= Fradato And StampDate <= Tildato Then
                .Cells(Rad, 1).Value = val(aTemp(i, 9))
                .Cells(Rad, 2).Value = aStaff(aTemp(i, 8) + 1, 2)
                If Not aTemp(i, 4) = "" Then
                    .Cells(Rad, 3).Value = StampDate
                    .Cells(Rad, 4).Value = TimeSerial( _
                        CInt(Mid(aTemp(i, 4), 2, 2)), _
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


' ============================================================
'  Hjelpefunksjoner
' ============================================================

Public Function Filvelger(Oppstart As String) As String
    Dim fd              As FileDialog
    Dim vrtSelectedItem As Variant
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .InitialFileName = Oppstart
        If .Show = -1 Then
            For Each vrtSelectedItem In .SelectedItems
                Filvelger = vrtSelectedItem
                Exit Function
            Next vrtSelectedItem
        End If
    End With
    Set fd = Nothing
End Function

Function FileExists(fname) As Boolean
    FileExists = (Dir(fname) <> "")
End Function

Public Function FindStaff(aRR As Variant, val As Variant) As Long
    Dim r As Long
    For r = LBound(aRR) To UBound(aRR)
        If aRR(r, 1) = val Then
            FindStaff = r
            Exit Function
        End If
    Next r
End Function

Public Function FindDate(aRR As Variant, val As Variant) As Long
    Dim r As Long
    For r = LBound(aRR) To UBound(aRR)
        If aRR(r, 1) = val Then
            FindDate = r
            Exit Function
        End If
    Next r
End Function

' Returnerer månedsnummer (to siffer) fra månedsnavn
Function MonthNumber(myMonthName As String) As String
    Dim sYear As String
    sYear = " " & Str(ThisWorkbook.Sheets("Settings").Cells(6, 2).Value)
    MonthNumber = Format(Month(DateValue("1 " & myMonthName & sYear)), "00")
End Function

Public Function LagStorBokstav(sWord As String) As String
    LagStorBokstav = StrConv(sWord, vbProperCase)
End Function

Public Function TallMedNull(lTall As Long, iAntall As Integer) As String
    TallMedNull = Format(lTall, String(iAntall, "0"))
End Function

' Returnerer Excel-kolonnebokstav fra kolonnenummer
Function Bokstav(Kolonne As Variant) As String
    Bokstav = Split(Cells(1, CLng(Kolonne)).Address, "$")(1)
End Function

' Returnerer ukedagnavn (engelsk) fra dato
Function GpDay(gDay As Date) As String
    Select Case Weekday(gDay)
        Case 2: GpDay = "Monday"
        Case 3: GpDay = "Tuesday"
        Case 4: GpDay = "Wednesday"
        Case 5: GpDay = "Thursday"
        Case 6: GpDay = "Friday"
        Case 7: GpDay = "Saturday"
        Case 1: GpDay = "Sunday"
    End Select
End Function

' Returnerer filnavn / månedsnavn for rapport
' FIX: bruker MonthName() i stedet for 12 If-setninger
Function GiveName(mNo As Integer, Nummer As Integer) As String
    Dim mName   As String
    Dim sYear   As String
    mName = MonthName(mNo)
    sYear = Str(ThisWorkbook.Sheets("Settings").Cells(5, 2).Value)
    If Nummer <> 2 Then
        GiveName = "Time Report " & mName & " " & sYear & ".xlsx"
    Else
        GiveName = "Time Report " & mName & " " & sYear
    End If
End Function

' Velger farge basert på kode
Public Function VelgFarge(Kode As String) As Long
    Select Case Right(Kode, 1)
        Case "0": VelgFarge = RGB(224, 224, 224)   ' Grå
        Case "1": VelgFarge = RGB(255, 255, 204)   ' Gul
        Case "2": VelgFarge = RGB(229, 255, 204)   ' Grønn
    End Select
    If Right(Kode, 3) = "ett" Then VelgFarge = RGB(204, 229, 255)   ' Blå
    If VelgFarge = 0 Then VelgFarge = RGB(255, 255, 255)             ' Hvit
End Function



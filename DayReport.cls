VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DayReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    ' FIX: riktig typedeklarasjon (var "Dim i, j As Long" → kun j ble Long)
    Dim i       As Long
    Dim j       As Long
    Dim Rad     As Long    ' FIX: var Integer – bør være Long for store ark
    Dim Kol     As Long
    Dim Navn    As String
    Dim KolNavn As String

    Rad    = Target.Row
    Kol    = Target.Column
    KolNavn = UCase(Cells(1, Kol).Value)

    Select Case KolNavn
        Case "NAME"
            ' Finn første ikke-tomme navn i kolonnen
            For i = 2 To 30
                If Not Cells(i, Kol).Value = "" Then
                    Navn = Cells(i, Kol).Value
                    Exit For
                End If
            Next i
            Cells(Rad, Kol).Value     = Navn
            Cells(Rad, Kol + 1).Value = ThisWorkbook.Sheets("Settings").Cells(15, 2).Value
            Cells(Rad, Kol + 2).Value = ThisWorkbook.Sheets("Settings").Cells(16, 2).Value

        Case "TIME IN"
            Cells(Rad, Kol).Value = ThisWorkbook.Sheets("Settings").Cells(15, 2).Value

        Case "TIME OUT"
            Cells(Rad, Kol).Value = ThisWorkbook.Sheets("Settings").Cells(16, 2).Value
    End Select
End Sub
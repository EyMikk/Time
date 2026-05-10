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
    Dim Adresse As Range
    Dim i, j As Long
    Dim Rad, Kol As Integer
    Dim Navn, KolNavn As String
    
    
    Rad = Target.Row
    Kol = Target.Column
    KolNavn = UCase(Cells(1, Kol).Value)

    If KolNavn = "NAME" Then

        For i = 2 To 30
            If Not Cells(i, Kol).Value = "" Then
                Navn = Cells(i, Kol).Value
                Exit For
            End If
        Next i
        Cells(Rad, Kol).Value = Navn
        Cells(Rad, Kol + 1).Value = ThisWorkbook.Sheets("Settings").Cells(15, 2).Value
        Cells(Rad, Kol + 2).Value = ThisWorkbook.Sheets("Settings").Cells(16, 2).Value
    End If

    If KolNavn = "TIME IN" Then
        Cells(Rad, Kol).Value = ThisWorkbook.Sheets("Settings").Cells(15, 2).Value
    End If

    If KolNavn = "TIME OUT" Then
        Cells(Rad, Kol).Value = ThisWorkbook.Sheets("Settings").Cells(16, 2).Value
    End If
End Sub

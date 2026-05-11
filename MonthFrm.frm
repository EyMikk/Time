VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MonthFrm 
   Caption         =   "Select Month"
   ClientHeight    =   3048
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   4605
   OleObjectBlob   =   "MonthFrm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "MonthFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdBoxOK_Click()
    ' FIX: riktig typedeklarasjon (var "Dim Fradato, Tildato As Date" → begge Variant)
    Dim Fradato As Date
    Dim Tildato As Date
    Dim Y       As Integer
    Dim M       As Integer
    Dim Y2      As Integer
    Dim M2      As Integer
    Dim Sett    As Worksheet

    Set Sett = ThisWorkbook.Sheets("Settings")

    Sett.Cells(5, 2).Value = ComboBox1.ListIndex
    M  = Sett.Cells(5, 2).Value - 1
    Y  = Sett.Cells(6, 2).Value
    Y2 = Y
    M2 = M + 1

    ' Håndter januar: gå til desember forrige år
    If M2 = 1 Then
        Y = Y2 - 1
        M = 12
    End If

    Fradato = DateSerial(Y, M, 25)
    Tildato = DateSerial(Y2, M2, 24)

    Unload Me

    Sett.Cells(7, 2).Value = Fradato
    Sett.Cells(8, 2).Value = Tildato

    MsgBox "This will report the time from " & Fradato & " to " & Tildato
End Sub

VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MonthFrm 
   Caption         =   "Select Month"
   ClientHeight    =   3045
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
Dim Fradato, Tildato As Date
Dim Y, M, D, Y2, M2 As Integer
Dim Sett As Worksheet
Dim tall As Integer
Set Sett = ThisWorkbook.Sheets("Settings")

ThisWorkbook.Sheets("Settings").Cells(5, 2).Value = ComboBox1.ListIndex
M = Sett.Cells(5, 2).Value - 1
Y = Sett.Cells(6, 2).Value
Y2 = Y
M2 = M + 1
If M2 = 1 Then Y = Y2 - 1
If M2 = 1 Then M = 12


Fradato = DateSerial(Y, M, 25)
Tildato = DateSerial(Y2, M2, 24)


'Tildato = Fradato + 30
Unload Me

Sett.Cells(7, 2).Value = Fradato
Sett.Cells(8, 2).Value = Tildato

MsgBox ("This will report the time from " & Fradato & " to " & Tildato)

End Sub



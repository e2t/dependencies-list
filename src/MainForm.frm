VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MainForm 
   Caption         =   "Список зависимостей модели"
   ClientHeight    =   9705.001
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   14625
   OleObjectBlob   =   "MainForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub allRad_Click()
    FillDeps
End Sub

Private Sub hieRad_Click()
    FillDeps
End Sub

Private Sub lstDeps_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If Shift = 2 And KeyCode = vbKeyC Then  '1 - Shift, 2 - Ctrl, 3 - Alt
        CopySelectedInClipboard GetSelected
    End If
End Sub

Private Sub topRad_Click()
    FillDeps
End Sub

Private Sub btnCancel_Click()
    ExitApp
End Sub

Private Sub btnOpen_Click()
    OpenDocs GetSelected
    ExitApp
End Sub

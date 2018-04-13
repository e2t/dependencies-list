VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MainForm 
   Caption         =   "Список зависимостей модели"
   ClientHeight    =   9705.001
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   14610
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

Private Sub UserForm_Initialize()
    '''The resize must be first!
    Me.Width = GetSystemMetrics(SM_CXSCREEN) - 600
    Me.Height = GetSystemMetrics(SM_CYSCREEN) - 400
    '''The resize must be first!
    
    Me.lstDeps.ColumnWidths = "2.2 cm;"
    Me.allRad.Caption = "Показать все"
    Me.topRad.Caption = "Только верхнего уровня"
End Sub

Private Sub UserForm_Resize()
    Me.allRad.Top = Me.Height - 45
    Me.topRad.Top = Me.allRad.Top
    Me.hieRad.Top = Me.allRad.Top

    Me.btnCancel.Top = Me.Height - 51
    Me.btnOpen.Top = Me.btnCancel.Top

    Me.btnCancel.Left = Me.Width - 81
    Me.btnOpen.Left = Me.btnCancel.Left - 78

    Me.lstDeps.Width = Me.Width - 15
    Me.lstDeps.Height = Me.Height - 61
End Sub

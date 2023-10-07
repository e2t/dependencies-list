VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MainForm 
   Caption         =   "DependenciesList 23.1"
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
    If Shift = 2 Then  '1 - Shift, 2 - Ctrl, 3 - Alt
        If KeyCode = vbKeyC Then
            CopySelectedInClipboard GetSelected
        ElseIf KeyCode = vbKeyA Then
            SelectAll
        End If
    End If
End Sub

Private Sub spnWidth_Change()
    Me.width = Me.spnWidth.value
End Sub

Private Sub spnHeight_Change()
    Me.height = Me.spnHeight.value
End Sub

Private Sub TextBoxFilter_Change()
    SelectFilterListbox Me.lstDeps, Me.TextBoxFilter.text
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

Private Sub TrySetSpinValue(spn As SpinButton, value As Integer)
    If value < spn.min Then
        spn.value = spn.min
    ElseIf value > spn.max Then
        spn.value = spn.max
    Else
        spn.value = value
    End If
End Sub

Private Sub UserForm_Initialize()
    Dim oldWidth As Integer
    Dim oldHeight As Integer
    Dim maxWidth As Integer
    Dim maxHeight As Integer
    
    maxWidth = MaximizedWidth
    maxHeight = MaximizedHeight
    oldWidth = GetIntSetting(WidthSetting, maxWidth)
    oldHeight = GetIntSetting(HeightSetting, maxHeight)
    
    Me.spnWidth.max = maxWidth
    Me.width = oldWidth
    Me.spnHeight.max = maxHeight
    Me.height = oldHeight
    
    '''It must be after resizing.
    Me.lstDeps.ColumnWidths = "60;"
    Me.hieRad.Caption = kStruct + " of this conf."
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    SaveIntSetting WidthSetting, Me.spnWidth.value
    SaveIntSetting HeightSetting, Me.spnHeight.value
End Sub

Private Sub UserForm_Resize()
    Me.allRad.Top = Me.height - 45
    Me.topRad.Top = Me.allRad.Top
    Me.hieRad.Top = Me.allRad.Top

    Me.btnCancel.Top = Me.height - 51
    Me.btnOpen.Top = Me.btnCancel.Top

    Me.btnCancel.Left = Me.width - 81
    Me.btnOpen.Left = Me.btnCancel.Left - 78

    Me.lstDeps.width = Me.width - 15
    Me.lstDeps.height = Me.height - 90
    
    TrySetSpinValue Me.spnWidth, Me.width
    Me.boxWidth.value = Me.spnWidth.value
    TrySetSpinValue Me.spnHeight, Me.height
    Me.boxHeight.value = Me.spnHeight.value
End Sub



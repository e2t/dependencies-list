Attribute VB_Name = "DrySettings"
Option Explicit

'Required:
'''
'Public Const MacroName = "MacroName"
'Public Const MacroSection = "Main"
'''

Sub SaveStrSetting(key As String, value As String)
    SaveSetting MacroName, MacroSection, key, value
End Sub

Sub SaveIntSetting(key As String, value As Integer)
    SaveSetting MacroName, MacroSection, key, Str(value)
End Sub

Sub SaveBoolSetting(key As String, value As Boolean)
    SaveSetting MacroName, MacroSection, key, Str(CInt(value))
End Sub

Function GetStrSetting(key As String, Optional fallbackValue As String = "") As String
    GetStrSetting = GetSetting(MacroName, MacroSection, key, fallbackValue)
End Function

Function GetIntSetting(key As String, Optional fallbackValue As Integer = 0) As Integer
    Dim value As String
    
    value = GetSetting(MacroName, MacroSection, key, "")
    If IsNumeric(value) Then
        GetIntSetting = CInt(value)
    Else
        GetIntSetting = fallbackValue
    End If
End Function

Function GetBoolSetting(key As String, Optional fallbackValue As Boolean = False) As Boolean
    Dim value As String
    
    value = GetSetting(MacroName, MacroSection, key, "")
    If IsNumeric(value) Then
        GetBoolSetting = CInt(value)
    Else
        GetBoolSetting = fallbackValue
    End If
End Function

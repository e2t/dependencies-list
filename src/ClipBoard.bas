Attribute VB_Name = "ClipBoard"
Option Explicit

'https://stackoverflow.com/questions/14219455/excel-vba-code-to-copy-a-specific-string-to-clipboard/60896244#60896244

Function CopyInClipboard$(Optional s$)
    Dim v: v = s  'Cast to variant for 64-bit VBA support
    With CreateObject("htmlfile")
    With .parentWindow.clipboardData
        Select Case True
            Case Len(s): .SetData "text", v
            Case Else:   CopyInClipboard = .GetData("text")
        End Select
    End With
    End With
End Function

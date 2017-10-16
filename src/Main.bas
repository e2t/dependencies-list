Attribute VB_Name = "Main"
Option Explicit

Const baseIndent As String = "|    "
Const labelSupressed As String = "*ПОГАШЕНО* "

Dim swApp As Object
Dim topDeps() As String
Dim allDeps() As String
Dim hieDeps As Collection
Dim depModels As Dictionary
Dim currentDocName As String
Dim currentDirName As String
    
Sub Main()
    Dim doc As ModelDoc2
    
    Set swApp = Application.SldWorks
    Set doc = swApp.ActiveDoc
    If doc Is Nothing Then Exit Sub
    
    Erase topDeps
    Erase allDeps
    Set hieDeps = Nothing
    Set depModels = New Dictionary
    currentDocName = doc.GetPathName
    currentDirName = Left(currentDocName, InStrRev(currentDocName, "\"))
    FillDeps
    MainForm.Show
End Sub

Function GetTopDeps() As String()
    If IsArrayEmpty(topDeps) Then
        topDeps = GetDeps(currentDocName, False)
    End If
    GetTopDeps = topDeps
End Function

Function GetAllDeps() As String()
    If IsArrayEmpty(allDeps) Then
        allDeps = GetDeps(currentDocName, True)
    End If
    GetAllDeps = allDeps
End Function

Function GetHieDeps() As Collection
    If hieDeps Is Nothing Then
        Set hieDeps = New Collection
        GetRecursiveDeps currentDocName, 0, hieDeps
    End If
    Set GetHieDeps = hieDeps
End Function

Sub GetRecursiveDeps(docname As String, level As Integer, ByRef deps As Collection)
    Dim i As Variant
    Dim depname As String
    Dim doc As ModelDoc2
    Dim indent As String
    
    indent = GetIndent(level)
    For Each i In GetDeps(docname, False)
        depname = i
        deps.Add indent & depname
        Set doc = depModels(depname)
        If Not doc Is Nothing Then
            If doc.GetType = swDocASSEMBLY Then
                GetRecursiveDeps depname, level + 1, deps
            End If
        End If
    Next
End Sub

Function GetIndent(level As Integer) As String
    Dim i As Integer
    
    GetIndent = ""
    For i = 1 To level
        GetIndent = GetIndent & baseIndent
    Next
End Function

Sub FillThisDeps(deps As Variant)
    Dim i As Variant
    
    MainForm.lstDeps.Clear
    For Each i In deps
        MainForm.lstDeps.AddItem i
    Next
End Sub

Function FillDeps()  'mask for button
    Dim inTitle As String
    
    If MainForm.allRad.Value Then
        FillThisDeps GetAllDeps
        inTitle = "все"
    ElseIf MainForm.topRad.Value Then
        FillThisDeps GetTopDeps
        inTitle = "только верхнего уровня"
    Else
        FillThisDeps GetHieDeps
        inTitle = "иерархично"
    End If
    MainForm.Caption = "Список зависимостей модели (" & inTitle & ")"
End Function

' TODO: use it
Function RelativePath(filename As String) As String
    If filename Like currentDirName & "*" Then
        RelativePath = Right(filename, Len(filename) - Len(currentDirName))
    Else
        RelativePath = filename
    End If
End Function

Function GetDeps(docname As String, resursive As Boolean) As String()
    Dim deps() As String
    Dim keys() As String
    Dim i As Integer
    Dim doc As ModelDoc2
    
    deps = swApp.GetDocumentDependencies2(docname, resursive, False, False)
    ReDim keys((UBound(deps) - LBound(deps) + 1) / 2 - 1)
    For i = LBound(deps) To UBound(deps) - 1 Step 2
        keys(i / 2) = deps(i + 1)
    Next
    BubbleSort keys
    For i = LBound(keys) To UBound(keys)
        Set doc = swApp.GetOpenDocumentByName(keys(i))
        If doc Is Nothing Then
            keys(i) = labelSupressed & keys(i)
        End If
        If Not depModels.Exists(keys(i)) Then
            depModels.Add keys(i), doc
        End If
    Next
    GetDeps = keys
End Function

Function GetSelected() As String()
    Dim i As Integer, j As Integer
    Dim selected() As String
    
    ReDim selected(MainForm.lstDeps.ListCount)
    j = 0
    For i = 0 To MainForm.lstDeps.ListCount - 1
        If MainForm.lstDeps.selected(i) Then
            selected(j) = GetCleanName(MainForm.lstDeps.List(i))  'it need only for hierarchy
            j = j + 1
        End If
    Next
    ReDim Preserve selected(j - 1)
    GetSelected = selected
End Function

Function GetCleanName(line As String) As String
    Dim ary() As String
    
    ary = Split(line, baseIndent)
    GetCleanName = ary(UBound(ary))
End Function

Sub OpenDocs(selected() As String)
    Dim i As Variant
    Dim key As String
    Dim doc As ModelDoc2
    
    For Each i In selected
        key = i
        Set doc = depModels(key)
        If Not doc Is Nothing Then
            ShowDoc doc.GetPathName
        End If
    Next
End Sub

Function ExitApp()  'mask for button
    Unload MainForm
    End
End Function

Sub BubbleSort(ByRef arr As Variant)
    Dim i As Integer
    Dim j As Integer
    Dim tmp As Variant
    
    For i = LBound(arr) To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If arr(i) > arr(j) Then
                tmp = arr(i)
                arr(i) = arr(j)
                arr(j) = tmp
            End If
        Next
    Next
End Sub

Function JoinStrings(lines() As String, sep As String) As String
    Dim i As Integer
    Dim lowIndex As Integer
    
    If IsArrayEmpty(lines) Then
        JoinStrings = ""
    Else
        lowIndex = LBound(lines)
        JoinStrings = lines(lowIndex)
        For i = lowIndex + 1 To UBound(lines)
            JoinStrings = JoinStrings & sep & lines(i)
        Next
    End If
End Function

Function IsArrayEmpty(anArray As Variant) As Boolean
    Dim i As Integer
  
    On Error Resume Next
        i = UBound(anArray, 1)
    If err.Number = 0 Then
        IsArrayEmpty = False
    Else
        IsArrayEmpty = True
    End If
End Function

Function ShowDoc(name As String) As ModelDoc2
    Dim err As swActivateDocError_e
    
    Set ShowDoc = swApp.ActivateDoc3(name, False, swDontRebuildActiveDoc, err)
End Function

Sub CopySelectedInClipboard(selected() As String)
    CopyInClipBoard JoinStrings(selected, vbNewLine)
End Sub

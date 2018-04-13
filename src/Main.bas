Attribute VB_Name = "Main"
Option Explicit

Public Declare PtrSafe Function GetSystemMetrics Lib "user32.dll" (ByVal index As Long) As Long
Public Const SM_CXSCREEN = 0
Public Const SM_CYSCREEN = 1

Const BASE_INDENT As String = "|    "
Const LABEL_NOT_LOADED As String = "НЕ ОТКРЫТ"
Const LABEL_NOT_FOUND As String = "НЕ НАЙДЕН"

Const COL_IS_LOADED As Integer = 0
Const COL_FULL_NAME As Integer = 1

Dim swApp As Object

Dim topDeps() As Depend_t
Dim allDeps() As Depend_t
Dim hieDeps As Collection
Public isAllDepsComputed As Boolean
Public isTopDepsComputed As Boolean
Public isHieDepsComputed As Boolean
Dim isAllDepsCountShowed As Boolean
Dim isTopDepsCountShowed As Boolean
Dim isHieDepsCountShowed As Boolean

Dim depModels As Dictionary  'key: fullName, value: ModelDoc2
Dim scriptFSO As FileSystemObject

Dim currentDocName As String
Dim currentDirName As String

Sub Main()
    Dim doc As ModelDoc2
    
    Set swApp = Application.SldWorks
    Set doc = swApp.ActiveDoc
    If doc Is Nothing Then Exit Sub
    
    Erase allDeps
    Erase topDeps
    Set hieDeps = Nothing
    isAllDepsComputed = False
    isTopDepsComputed = False
    isHieDepsComputed = False
    isAllDepsCountShowed = False
    isTopDepsCountShowed = False
    isHieDepsCountShowed = False
    
    Set depModels = New Dictionary
    Set scriptFSO = New FileSystemObject

    currentDocName = doc.GetPathName
    currentDirName = Left(currentDocName, InStrRev(currentDocName, "\"))
    
    FillDeps
    MainForm.Show
End Sub

Function GetAllDeps() As Depend_t()
    If Not isAllDepsComputed Then
        allDeps = GetDeps(currentDocName, True)
        isAllDepsComputed = True
    End If
    GetAllDeps = allDeps
End Function

Function GetTopDeps() As Depend_t()
    If Not isTopDepsComputed Then
        topDeps = GetDeps(currentDocName, False)
        isTopDepsComputed = True
    End If
    GetTopDeps = topDeps
End Function

Function GetHieDeps() As Collection
    If Not isHieDepsComputed Then
        Set hieDeps = New Collection
        GetRecursiveDeps currentDocName, 0, hieDeps
        isHieDepsComputed = True
    End If
    Set GetHieDeps = hieDeps
End Function

Function CheckStatusModel(doc As ModelDoc2, filename As String) As FileStatus_t
    If Not doc Is Nothing Then
        CheckStatusModel = IS_LOADED
    ElseIf scriptFSO.FileExists(filename) Then
        CheckStatusModel = IS_NOT_LOADED
    Else
        CheckStatusModel = IS_NOT_FOUND
    End If
End Function

Sub GetRecursiveDeps(docname As String, level As Integer, ByRef deps As Collection)
    Dim i As Variant
    Dim depname As String
    Dim doc As ModelDoc2
    Dim indent As String
    Dim depend As Depend_t
    Dim get_deps As Variant
    
    indent = GetIndent(level)
    get_deps = GetDeps(docname, False)
    If IsArrayEmpty(get_deps) Then
        Exit Sub
    End If
    For Each i In get_deps
        depname = i.fullName
        Set doc = depModels(depname)
        Set depend = New Depend_t
        depend.fullName = indent & depname
        depend.status = CheckStatusModel(doc, depname)
        deps.Add depend
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
        GetIndent = GetIndent & BASE_INDENT
    Next
End Function

Sub FillThisDeps(deps As Variant)
    Dim i As Variant
    Dim depend As Depend_t
    
    With MainForm.lstDeps
        .Clear
        For Each i In deps
            Set depend = i
            .AddItem
            Select Case depend.status
                Case IS_NOT_LOADED
                    .List(.ListCount - 1, COL_IS_LOADED) = LABEL_NOT_LOADED
                Case IS_NOT_FOUND
                    .List(.ListCount - 1, COL_IS_LOADED) = LABEL_NOT_FOUND
                Case Else
                    .List(.ListCount - 1, COL_IS_LOADED) = ""
            End Select
            .List(.ListCount - 1, COL_FULL_NAME) = depend.fullName
        Next
    End With
End Sub

Function ArraySize(arr As Variant) As Integer
    If IsArrayEmpty(arr) Then
        ArraySize = 0
    Else
        ArraySize = UBound(arr) - LBound(arr) + 1
    End If
End Function

Function FillDeps()  'mask for button
    Dim inTitle As String
    
    With MainForm
        If .allRad.Value Then
            FillThisDeps GetAllDeps
            inTitle = "все"
            If Not isAllDepsCountShowed Then
                .allRad.Caption = "Показать все (" & ArraySize(allDeps) & ")"
                isAllDepsCountShowed = True
            End If
        ElseIf .topRad.Value Then
            FillThisDeps GetTopDeps
            inTitle = "только верхнего уровня"
            If Not isTopDepsCountShowed Then
                .topRad.Caption = "Только верхнего уровня (" & ArraySize(topDeps) & ")"
                isTopDepsCountShowed = True
            End If
        Else
            FillThisDeps GetHieDeps
            inTitle = "иерархично"
        End If
        .Caption = "Список зависимостей модели (" & inTitle & ")"
    End With
End Function

' TODO: use it
Function RelativePath(filename As String) As String
    If filename Like currentDirName & "*" Then
        RelativePath = Right(filename, Len(filename) - Len(currentDirName))
    Else
        RelativePath = filename
    End If
End Function

Function GetDeps(docname As String, resursive As Boolean) As Depend_t()
    Dim result() As Depend_t
    Dim deps As Variant  'String() or empty
    Dim i As Integer
    Dim doc As ModelDoc2
    
    deps = swApp.GetDocumentDependencies2(docname, resursive, False, False)
    If IsArrayEmpty(deps) Then
        Exit Function
    End If
    ReDim result((UBound(deps) - LBound(deps) + 1) / 2 - 1)
    For i = LBound(deps) To UBound(deps) - 1 Step 2
        Set result(i / 2) = New Depend_t
        result(i / 2).fullName = deps(i + 1)
    Next
    BubbleSortDepends result
    For i = LBound(result) To UBound(result)
        Set doc = swApp.GetOpenDocumentByName(result(i).fullName)
        result(i).status = CheckStatusModel(doc, result(i).fullName)
        If Not depModels.Exists(result(i).fullName) Then
            depModels.Add result(i).fullName, doc
        End If
    Next
    GetDeps = result
End Function

Function GetSelected() As String()
    Dim i As Integer, j As Integer
    Dim selected() As String
    
    ReDim selected(MainForm.lstDeps.ListCount)
    j = -1
    For i = 0 To MainForm.lstDeps.ListCount - 1
        If MainForm.lstDeps.selected(i) Then
            j = j + 1
            selected(j) = GetCleanName(MainForm.lstDeps.List(i, COL_FULL_NAME)) 'it need only for hierarchy
        End If
    Next
    If j >= 0 Then
        ReDim Preserve selected(j)
        GetSelected = selected
    End If
End Function

Function GetCleanName(line As String) As String
    Dim ary() As String
    
    ary = Split(line, BASE_INDENT)
    GetCleanName = ary(UBound(ary))
End Function

Sub OpenDocs(selected() As String)
    Dim i As Variant
    Dim key As String
    Dim doc As ModelDoc2
    
    If IsArrayEmpty(selected) Then
        Exit Sub
    End If
    
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

Sub SwapObj(ByRef first As Object, ByRef second As Object)
    Dim tmp As Object
    
    Set tmp = first
    Set first = second
    Set second = tmp
End Sub

Sub BubbleSortDepends(ByRef arr() As Depend_t)
    Dim i As Integer
    Dim j As Integer
    Dim need_sorted As Boolean
    Dim penult_index As Integer
    Dim tmp As Depend_t
    
    penult_index = UBound(arr) - 1
    need_sorted = True
    i = LBound(arr)
    While (i <= penult_index) And need_sorted
        need_sorted = False
        For j = LBound(arr) To penult_index - i
            If arr(j).fullName > arr(j + 1).fullName Then
                SwapObj arr(j), arr(j + 1)
                need_sorted = True
            End If
        Next
        i = i + 1
    Wend
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

Function IsArrayEmpty(ByRef anArray As Variant) As Boolean
    Dim i As Integer
  
    On Error GoTo ArrayIsEmpty
    IsArrayEmpty = LBound(anArray) > UBound(anArray)
    Exit Function
ArrayIsEmpty:
    IsArrayEmpty = True
End Function

Function ShowDoc(name As String) As ModelDoc2
    Dim err As swActivateDocError_e
    
    Set ShowDoc = swApp.ActivateDoc3(name, False, swDontRebuildActiveDoc, err)
End Function

Sub CopySelectedInClipboard(selected() As String)
    CopyInClipBoard JoinStrings(selected, vbNewLine)
End Sub

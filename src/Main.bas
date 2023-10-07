Attribute VB_Name = "Main"
Option Explicit

Public Const MacroName = "DependenciesList"
Public Const MacroSection = "Main"
Public Const WidthSetting = "WidthScreen"
Public Const HeightSetting = "HeightScreen"

Const BASE_INDENT As String = "|    "
Const LABEL_NOT_LOADED As String = "not open"
Const LABEL_NOT_FOUND As String = "NOT FOUND"
Public Const kEverything = "Everything"
Public Const kTopOnly = "Top level only"
Public Const kStruct = "Structure"

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

Public currentDocName As String
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

Function CheckStatusModel(doc As ModelDoc2, fileName As String) As FileStatus_t
    If Not doc Is Nothing Then
        CheckStatusModel = IS_LOADED
    ElseIf scriptFSO.FileExists(fileName) Then
        CheckStatusModel = IS_NOT_LOADED
    Else
        CheckStatusModel = IS_NOT_FOUND
    End If
End Function

Sub GetRecursiveDeps(docname As String, level As Integer, ByRef deps As Collection)
    Dim I As Variant
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
    For Each I In get_deps
        depname = I.fullName
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
    Dim I As Integer
    
    GetIndent = ""
    For I = 1 To level
        GetIndent = GetIndent & BASE_INDENT
    Next
End Function

Sub FillThisDeps(deps As Variant)
    Dim I As Variant
    Dim depend As Depend_t
    
    With MainForm.lstDeps
        .Clear
        For Each I In deps
            Set depend = I
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
        If .allRad.value Then
            FillThisDeps GetAllDeps
            inTitle = LCase(kEverything)
            If Not isAllDepsCountShowed Then
                .allRad.Caption = kEverything & " (" & ArraySize(allDeps) & ")"
                isAllDepsCountShowed = True
            End If
        ElseIf .topRad.value Then
            FillThisDeps GetTopDeps
            inTitle = LCase(kTopOnly)
            If Not isTopDepsCountShowed Then
                .topRad.Caption = kTopOnly & " (" & ArraySize(topDeps) & ")"
                isTopDepsCountShowed = True
            End If
        Else
            FillThisDeps GetHieDeps
            inTitle = LCase(kStruct)
        End If
        SelectFilterListbox .lstDeps, .TextBoxFilter.text
    End With
End Function

' TODO: use it
Function RelativePath(fileName As String) As String
    If fileName Like currentDirName & "*" Then
        RelativePath = Right(fileName, Len(fileName) - Len(currentDirName))
    Else
        RelativePath = fileName
    End If
End Function

Function GetDeps(docname As String, resursive As Boolean) As Depend_t()
    Dim result() As Depend_t
    Dim deps As Variant  'String() or empty
    Dim I As Integer
    Dim doc As ModelDoc2
    
    deps = swApp.GetDocumentDependencies2(docname, resursive, False, False)
    If IsArrayEmpty(deps) Then
        Exit Function
    End If
    ReDim result((UBound(deps) - LBound(deps) + 1) / 2 - 1)
    For I = LBound(deps) To UBound(deps) - 1 Step 2
        Set result(I / 2) = New Depend_t
        result(I / 2).fullName = deps(I + 1)
    Next
    BubbleSortDepends result
    For I = LBound(result) To UBound(result)
        Set doc = swApp.GetOpenDocumentByName(result(I).fullName)
        result(I).status = CheckStatusModel(doc, result(I).fullName)
        If Not depModels.Exists(result(I).fullName) Then
            depModels.Add result(I).fullName, doc
        End If
    Next
    GetDeps = result
End Function

Function GetSelected() As String()
    Dim I As Integer, j As Integer
    Dim Selected() As String
    
    ReDim Selected(MainForm.lstDeps.ListCount)
    j = -1
    For I = 0 To MainForm.lstDeps.ListCount - 1
        If MainForm.lstDeps.Selected(I) Then
            j = j + 1
            Selected(j) = GetCleanName(MainForm.lstDeps.List(I, COL_FULL_NAME)) 'it need only for hierarchy
        End If
    Next
    If j >= 0 Then
        ReDim Preserve Selected(j)
        GetSelected = Selected
    End If
End Function

Function GetCleanName(line As String) As String
    Dim ary() As String
    
    ary = Split(line, BASE_INDENT)
    GetCleanName = ary(UBound(ary))
End Function

Sub OpenDocs(Selected() As String)
    Dim I As Variant
    Dim key As String
    Dim doc As ModelDoc2
    
    If IsArrayEmpty(Selected) Then
        Exit Sub
    End If
    
    For Each I In Selected
        key = I
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
    Dim I As Integer
    Dim j As Integer
    Dim need_sorted As Boolean
    Dim penult_index As Integer
    Dim tmp As Depend_t
    
    penult_index = UBound(arr) - 1
    need_sorted = True
    I = LBound(arr)
    While (I <= penult_index) And need_sorted
        need_sorted = False
        For j = LBound(arr) To penult_index - I
            If arr(j).fullName > arr(j + 1).fullName Then
                SwapObj arr(j), arr(j + 1)
                need_sorted = True
            End If
        Next
        I = I + 1
    Wend
End Sub

Function JoinStrings(lines() As String, sep As String) As String
    Dim I As Integer
    Dim lowIndex As Integer
    
    If IsArrayEmpty(lines) Then
        JoinStrings = ""
    Else
        lowIndex = LBound(lines)
        JoinStrings = lines(lowIndex)
        For I = lowIndex + 1 To UBound(lines)
            JoinStrings = JoinStrings & sep & lines(I)
        Next
    End If
End Function

Function IsArrayEmpty(ByRef anArray As Variant) As Boolean
    Dim I As Integer
  
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

Sub CopySelectedInClipboard(Selected() As String)
    CopyInClipboard JoinStrings(Selected, vbNewLine)
End Sub

Function SelectAll(Optional IsSelected As Boolean = True) 'hide
    Dim I As Variant
    
    With MainForm.lstDeps
        For I = 0 To .ListCount - 1
            .Selected(I) = IsSelected
        Next
    End With
End Function

Function DeselectAll() 'hide
    SelectAll False
End Function

Sub SelectFilterListbox(Lst As ListBox, text As String)
    Dim I As Integer
    
    If text = "" Then
        DeselectAll
    Else
        With MainForm.lstDeps
            For I = 0 To Lst.ListCount - 1
                .Selected(I) = InStr(1, .List(I, 1), text, vbTextCompare)
            Next
        End With
    End If
End Sub

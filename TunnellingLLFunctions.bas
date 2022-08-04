Attribute VB_Name = "TunnellingLLFunctions"
Function getWordApp() As Word.Application 'Dim getWordApp as word.app
    Dim wAppLocal As Word.Application
    On Error Resume Next
    Set wAppLocal = GetObject(, "Word.Application") ' try to get the word app if it is open
    If wAppLocal Is Nothing Then 'if word is not open
        Set wAppLocal = CreateObject("Word.Application")
    End If
    Set getWordApp = wAppLocal
End Function

Function UniqueVals(Col As Variant, Optional SheetName As String = vbNullString) As Variant
    'Return a 1-based array of the unique values in column Col
    Dim D As Variant, A As Variant, v As Variant
    Dim i As Long, n As Long, k As Long
    Dim ws As Worksheet
    Dim subRange As Range
    Dim rng As Range
    If Len(SheetName) = 0 Then
        'Set ws = ActiveSheet
        'Set ws = Application.ThisWorkbook.Sheets("CYP D&C Hazard Log")
        MsgBox "Error, no sheetname specified to obtain unique values"
    Else
        Set ws = Application.ThisWorkbook.Sheets(SheetName)
    End If
    k = 0
    lastRow = ws.UsedRange.Rows(ws.UsedRange.Rows.Count).Row
    On Error Resume Next
    Set subRange = ws.Range(Cells(3, Col), Cells(lastRow, Col)).SpecialCells(xlCellTypeVisible)
    lcount = ws.AutoFilter.Range.Columns(Col).SpecialCells(xlCellTypeVisible).Cells.Count
    ReDim A(1 To lcount)
    Set D = CreateObject("Scripting.Dictionary")
    For Each rng In subRange
        v = rng.Value
        If v = vbNullString Then
            'do nothing
        Else
            If Not D.Exists(v) Then
                D.Add v, 0
                k = k + 1
                A(k) = v
            End If
        End If
    Next rng
    If k = 0 Then
       ' A = Empty
    Else
        ReDim Preserve A(1 To k)
    End If
    UniqueVals = A

End Function


Sub word_Options(wordApp As Word.Application, toggle As Boolean) 'toggle = false, off, toggle = true, on
    
    wordApp.ScreenUpdating = toggle
    wordApp.Options.CheckGrammarAsYouType = toggle
    wordApp.Options.CheckGrammarWithSpelling = toggle
    wordApp.Options.CheckSpellingAsYouType = toggle
    wordApp.Options.AnimateScreenMovements = toggle
    wordApp.Options.BackgroundSave = toggle
    wordApp.Options.CheckHangulEndings = toggle
    wordApp.Options.DisableFeaturesbyDefault = Not toggle
    
End Sub

Function toggleParameters(toggle As String)
         ' Set excel parameters to speed up code execution
    If toggle = "On" Then
        Application.DisplayAlerts = False
        Application.Calculation = xlCalculationManual
        Application.ScreenUpdating = False
        Application.EnableEvents = False
        Application.AskToUpdateLinks = False
    Else
        Application.DisplayAlerts = True
        Application.Calculation = xlCalculationAuto
        Application.ScreenUpdating = True
        Application.EnableEvents = True
        Application.AskToUpdateLinks = True
    End If
End Function

Function clearFilters() As Boolean
    clearFilters = False
    Dim wkb As Workbook
    Set wkb = Application.ThisWorkbook
    Dim xWs As Worksheet
    On Error Resume Next
    For Each xWs In wkb.Worksheets
        xWs.ShowAllData
    Next
    clearFilters = True
End Function

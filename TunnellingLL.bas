Attribute VB_Name = "TunnellingLL"
Sub Tunnelling_LL()

    'new text here - testing git push etc'
    'this is another line'


    'Define key variables for use in the code
    ' Excel variables
    Dim wkb As Workbook, ws As Worksheet, wkbLL As Workbook, wsLL As Worksheet
    ' Word variables
    Dim wApp As Word.Application, wDoc As Word.Document, astrHeadings As Variant

    startTime = Timer ' timing the execution of the program
    ' Set excel parameters to speed up code execution
    toggleParameters ("On")
    
    ' Get LL spreadsheet
    MsgBox "Select the Lessons Learnt Spreadsheet to append to"
    fileToOpen = Application.GetOpenFilename
    Application.Workbooks.Open (fileToOpen)
    Set wkbLL = Application.ActiveWorkbook
    Set wsLL = wkbLL.Sheets("ICI Register")
    
    ' Get column headings
    Set headingRange = wsLL.Range("5:5")
    xStr = "Title"
    Set llTitle = headingRange.Find(xStr, , xlValues, xlWhole, , , True)
    xStr = "Description"
    Set llDesc = headingRange.Find(xStr, , xlValues, xlWhole, , , True)
    xStr = "Date (start)"
    Set llDS = headingRange.Find(xStr, , xlValues, xlWhole, , , True)
    xStr = "Date (completion)"
    Set llDC = headingRange.Find(xStr, , xlValues, xlWhole, , , True)
    xStr = "Project"
    Set llProject = headingRange.Find(xStr, , xlValues, xlWhole, , , True)
    xStr = "Item"
    Set llItem = headingRange.Find(xStr, , xlValues, xlWhole, , , True)
    xStr = "Area"
    Set llArea = headingRange.Find(xStr, , xlValues, xlWhole, , , True)
    xStr = "Category"
    Set llCat = headingRange.Find(xStr, , xlValues, xlWhole, , , True)
    xStr = "Benefits"
    Set llBen = headingRange.Find(xStr, , xlValues, xlWhole, , , True)
    xStr = "Source Document"
    Set llSD = headingRange.Find(xStr, , xlValues, xlWhole, , , True)
    xStr = "Section No"
    Set llSN = headingRange.Find(xStr, , xlValues, xlWhole, , , True)
    xStr = "Section Title"
    Set llST = headingRange.Find(xStr, , xlValues, xlWhole, , , True)
    xStr = "Keywords"
    Set llKW = headingRange.Find(xStr, , xlValues, xlWhole, , , True)
    
    ' Get LL Word Document
    Set wApp = getWordApp()
    
    MsgBox "Select the Word Document to consolidate Lessons Learnt from"
    fileToOpen = Application.GetOpenFilename
    wApp.Documents.Open (fileToOpen)
    Set wDoc = wApp.ActiveDocument
    
OpenAlready: 'if word is open and the document is open, document has already been set.
    wApp.Visible = True
    wDoc.Activate
    Debug.Print wDoc.Name

    ' Options for Word
    'word_Options wApp, False ' This is a function, false turns them off

    '-------------------- Main Loop to input data into Tables -----------------------------------------------
    Dim regEx As Object, matchCollection As Object, extractedString As String, currSHE As String, adoptCons() As String, openCons() As String, rejCons() As String
    Dim currRange As Word.Range, nextRange As Word.Range, extendRange As Word.Range
    
    'Regular expression to find the SHE in the headings
    Set regEx = CreateObject("VBScript.RegExp")
    With regEx
      .IgnoreCase = True
      .Global = False    ' Only look for 1 match; False is actually the default.
      .Pattern = "Lessons Learned"
    End With
    ' Additional variables for different uses
    p = 1 ' counter of what paragraph we are up to

    ' Setting parameters for the main loop
    currentLL = vbNullString
    matchLL = 0
    matchType = 0
    For Each Paragraph In wDoc.TablesOfContents(1).Range.Paragraphs 'searching paragraphs in table of contents
        'Paragraph.Range.Select
        headings = Split(Paragraph, Chr(9)) 'splits based on 'tab' character
        matchType = 0
        'On Error Resume Next
        Set matchCollection = regEx.Execute(headings(1))
        If matchCollection.Count = 0 Then
            currentLL = vbNullString
        Else
            currentLL = matchCollection(0)
        End If
        If currentLL = vbNullString Then 'matchLL indicates whether we have a 'Lessons Learnt' in the title or not
            matchLL = 0
        Else 'Filter and collect all hazards, causes, controls related to the SHE for the station or tunnel
            matchLL = 1
        End If

        currHeading = vbNullString
        currHeadingNo = vbNullString
        If matchType > 0 Then
            ' -- Follow the link to the section, navigate to the next table, and clean up the data to make a blank table
            wDoc.TablesOfContents(1).Range.Hyperlinks(p + 1).Follow NewWindow:=False, AddHistory:=True
            Set nextRange = wApp.Selection.Range
            wDoc.TablesOfContents(1).Range.Hyperlinks(p).Follow NewWindow:=False, AddHistory:=True
            ' Extend the range and check if table exists
            Set currRange = wApp.Selection.Range
            Set extendRange = wDoc.Range(currRange.Start, nextRange.Start)
            ' Initialising the heading values to the parent heading
            currHeading = extendRange.Paragraphs(1).Range.Text
            currHeadingNo = extendRange.Paragraphs(1).Range.ListFormat.ListString
            '--------------------------
            j = wsLL.UsedRange.Rows(wsLL.UsedRange.Rows.Count).Row + 1 'j finds the last entry in the LL register, and begins from that position + 1
            For i = 2 To extendRange.Paragraphs.Count
                varLength = extendRange.Paragraphs(i + 1)
                lenghtLL = Len(varLength)
                'Debug.Print extendRange.Paragraphs(i).Style
                If extendRange.Paragraphs(i).Style = "Heading 3 Numbered" And lengthLL > 10 Then
                    currHeading = extendRange.Paragraphs(i).Range.Text
                    currHeadingNo = extendRange.Paragraphs(i).Range.ListFormat.ListString
                    GoTo skipOutput
                End If
                'Debug.Print
                wsLL.Cells(j, llItem.Column).Value = CInt(wsLL.Cells(j - 1, llItem.Column).Value) + 1
                wsLL.Cells(j, llSD.Column).Value = wDoc.Name
                wsLL.Cells(j, llArea.Column).Value = "<User Input Required>"
                wsLL.Cells(j, llCat.Column).Value = "<User Input Required>"
                wsLL.Cells(j, llBen.Column).Value = "<User Input Required>"
                wsLL.Cells(j, llProject.Column).Value = "<User Input Required>"
                wsLL.Cells(j, llTitle.Column).Value = "<User Input Required>"
                wsLL.Cells(j, llSN.Column).Value = currHeadingNo
                wsLL.Cells(j, llST.Column).Value = currHeading
                wsLL.Cells(j, llDesc.Column).Value = extendRange.Paragraphs(i).Range.Text
                j = j + 1
skipOutput:
            Next i
        End If
        p = p + 1
    Next Paragraph
' --- END OF MAIN LOOP FOR ITERATING THROUGH DOCUMENT ---
    Dim minutesElapsed As Double
    Dim secondsElapsed As Double
    
    secondsElapsed = Round(Timer - startTime, 2)
    minutesElapsed = secondsElapsed / 60
    Debug.Print minutesElapsed
    'word_Options wApp, True ' this turns word parameters back on
    '---------------------------------------------------------------------------
       
    ' Reset excel parameters back to original state
    toggleParameters ("Off")

End Sub


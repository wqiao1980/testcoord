Attribute VB_Name = "MeshCoordinateUpdater"
Option Explicit

Public MeshCoordinateUpdateIsRunning As Boolean

Private Const TARGET_SHEET As String = "ARENA_MOD_PLWithDBM2-10"
Private Const SPRING_SHEET As String = "Spring2-10"
Private Const SPRING_DISTANCE_SHEET As String = "SpringDistance2-10"
Private Const SHEET10_SHEET As String = "Sheet10"
Private Const CURVE_FIRST_ROW As Long = 73
Private Const OUTPUT_FIRST_COL As Long = 26  ' Z
Private Const DEFAULT_OUTPUT_FIRST_ROW As Long = 10
Private Const SPRING_OUTPUT_FIRST_ROW As Long = 5
Private Const SPRING_DISTANCE_FIRST_ROW As Long = 9
Private Const SHEET10_FIRST_ROW As Long = 9
Private Const SPRING_TEXT_FILE_NAME As String = _
    "BPTiber_With_BM_FL6_TestBase_05_Lateral_Springs.txt"
Private Const INITIAL_DISPLACEMENT_TEXT_FILE_NAME As String = _
    "BPTiber_With_BM_FL6_TestBase_05_Initial_1P5m.txt"
Private Const TENSION_ADJUSTMENT_TEXT_FILE_NAME As String = _
    "BPTiber_With_BM_FL6_TestBase_05_TensionAdj.txt"
Private Const WORKFLOW_ALL As Long = 0
Private Const WORKFLOW_COORDINATES_ONLY As Long = 1
Private Const WORKFLOW_SPRING2 As Long = 2
Private Const WORKFLOW_SPRING_DISTANCE As Long = 3
Private Const WORKFLOW_SHEET10 As Long = 4
Private Const PI_VALUE As Double = 3.14159265358979

Public Sub RunAllWorkflows()
    UpdateGeneratedCoordinatesCore True, WORKFLOW_ALL
End Sub

Public Sub RebuildGeneratedCoordinates()
    RunAllWorkflows
End Sub

Public Sub RebuildCoordinatesOnly()
    UpdateGeneratedCoordinatesCore True, WORKFLOW_COORDINATES_ONLY
End Sub

Public Sub RunSpring2Workflow()
    UpdateGeneratedCoordinatesCore True, WORKFLOW_SPRING2
End Sub

Public Sub RunSpringDistanceWorkflow()
    UpdateGeneratedCoordinatesCore True, WORKFLOW_SPRING_DISTANCE
End Sub

Public Sub RunSheet10Workflow()
    UpdateGeneratedCoordinatesCore True, WORKFLOW_SHEET10
End Sub

Public Sub UpdateGeneratedCoordinatesCore( _
           ByVal ShowCompletionMessage As Boolean, _
           Optional ByVal WorkflowMode As Long = WORKFLOW_ALL)
    Dim ws As Worksheet
    Dim oldEnableEvents As Boolean
    Dim oldScreenUpdating As Boolean
    Dim anchorNodes() As Long
    Dim anchorX() As Double
    Dim anchorY() As Double
    Dim anchorZ() As Double
    Dim anchorCount As Long
    Dim curves As Collection
    Dim outputValues() As Variant
    Dim outputStartRow As Long
    Dim outputLastRow As Long
    Dim previousOutputLastRow As Long
    Dim generatedCount As Long
    Dim completionText As String

    If MeshCoordinateUpdateIsRunning Then Exit Sub
    MeshCoordinateUpdateIsRunning = True

    oldEnableEvents = Application.EnableEvents
    oldScreenUpdating = Application.ScreenUpdating
    Application.EnableEvents = False
    Application.ScreenUpdating = False

    On Error GoTo CleanFail

    If WorkflowMode < WORKFLOW_ALL Or WorkflowMode > WORKFLOW_SHEET10 Then
        Err.Raise vbObjectError + 2030, , "Unknown workflow selection."
    End If

    Set ws = ThisWorkbook.Worksheets(TARGET_SHEET)

    ' G:J normally mirrors A:D. Calculating this compact range first lets the
    ' macro use either a pasted G:J mesh or the formulas linked to A:D.
    ws.Range("G1:J" & CStr(CURVE_FIRST_ROW - 1)).Calculate

    LoadAnchors ws, anchorNodes, anchorX, anchorY, anchorZ, anchorCount
    If anchorCount < 2 Then
        Err.Raise vbObjectError + 2001, , _
            "At least two complete mesh anchor rows are required in A:D or G:J."
    End If

    Set curves = LoadCurves(ws, anchorNodes, anchorCount)

    previousOutputLastRow = FindLastOutputRow(ws)

    BuildOutputArray ws, anchorNodes, anchorX, anchorY, anchorZ, anchorCount, _
                     curves, outputValues, outputStartRow, outputLastRow, generatedCount

    ExtendOutputFormatting ws, outputStartRow, outputLastRow, previousOutputLastRow

    ws.Range(ws.Cells(outputStartRow, OUTPUT_FIRST_COL), _
             ws.Cells(outputLastRow, OUTPUT_FIRST_COL + 3)).Value2 = outputValues

    If previousOutputLastRow > outputLastRow Then
        ws.Range(ws.Cells(outputLastRow + 1, OUTPUT_FIRST_COL), _
                 ws.Cells(previousOutputLastRow, OUTPUT_FIRST_COL + 3)).ClearContents
    End If

    Select Case WorkflowMode
        Case WORKFLOW_ALL
            UpdateSpringSheet ws, outputStartRow, outputLastRow, generatedCount
            UpdateSpringDistanceSheet ws, outputStartRow, outputLastRow, generatedCount
            UpdateSheet10 ws, outputStartRow, outputLastRow, generatedCount
            completionText = _
                "Coordinates, Spring2-10, SpringDistance2-10, Sheet10, " & _
                "and all three text files were updated."

        Case WORKFLOW_COORDINATES_ONLY
            completionText = "Coordinates Z:AC were rebuilt."

        Case WORKFLOW_SPRING2
            UpdateSpringSheet ws, outputStartRow, outputLastRow, generatedCount
            completionText = _
                "Coordinates, Spring2-10, and the lateral-spring text file were updated."

        Case WORKFLOW_SPRING_DISTANCE
            UpdateSpringDistanceSheet ws, outputStartRow, outputLastRow, generatedCount
            completionText = _
                "Coordinates, SpringDistance2-10, and the initial-displacement text file were updated."

        Case WORKFLOW_SHEET10
            UpdateSheet10 ws, outputStartRow, outputLastRow, generatedCount
            completionText = _
                "Coordinates, Sheet10, and the tension-adjustment text file were updated."
    End Select

    If ShowCompletionMessage Then
        MsgBox completionText & vbCrLf & vbCrLf & _
               Format$(generatedCount, "#,##0") & _
               " coordinates were generated using " & _
               CStr(anchorCount) & " anchors and " & _
               CStr(curves.Count) & " curve definitions.", _
               vbInformation, "Manual workflow completed"
    End If

CleanExit:
    Application.EnableEvents = oldEnableEvents
    Application.ScreenUpdating = oldScreenUpdating
    MeshCoordinateUpdateIsRunning = False
    Exit Sub

CleanFail:
    MsgBox "The selected workflow was not completed: " & Err.Description, _
           vbExclamation, "Manual workflow"
    Resume CleanExit
End Sub

Private Sub LoadAnchors(ByVal ws As Worksheet, _
                        ByRef nodes() As Long, _
                        ByRef xValues() As Double, _
                        ByRef yValues() As Double, _
                        ByRef zValues() As Double, _
                        ByRef anchorCount As Long)
    Dim anchorMap As Object
    Dim lastAnchorRow As Long
    Dim rowNumber As Long
    Dim baseColumn As Long
    Dim nodeNumber As Long
    Dim key As String
    Dim keys As Variant
    Dim point As Variant
    Dim index As Long

    Set anchorMap = CreateObject("Scripting.Dictionary")
    anchorMap.CompareMode = vbTextCompare
    lastAnchorRow = FindAnchorBlockLastRow(ws)

    For rowNumber = 1 To lastAnchorRow
        ' Prefer G:J because the existing workbook calculations use those
        ' columns. Fall back to A:D when G:J is not a complete numeric row.
        If IsCompleteAnchorRow(ws, rowNumber, 7) Then
            baseColumn = 7
        ElseIf IsCompleteAnchorRow(ws, rowNumber, 1) Then
            baseColumn = 1
        Else
            baseColumn = 0
        End If

        If baseColumn > 0 Then
            nodeNumber = CLng(ws.Cells(rowNumber, baseColumn).Value2)
            key = CStr(nodeNumber)
            anchorMap(key) = Array( _
                CDbl(ws.Cells(rowNumber, baseColumn + 1).Value2), _
                CDbl(ws.Cells(rowNumber, baseColumn + 2).Value2), _
                CDbl(ws.Cells(rowNumber, baseColumn + 3).Value2))
        End If
    Next rowNumber

    anchorCount = anchorMap.Count
    If anchorCount = 0 Then Exit Sub

    ReDim nodes(0 To anchorCount - 1)
    ReDim xValues(0 To anchorCount - 1)
    ReDim yValues(0 To anchorCount - 1)
    ReDim zValues(0 To anchorCount - 1)

    keys = anchorMap.Keys
    For index = 0 To anchorCount - 1
        nodes(index) = CLng(keys(index))
        point = anchorMap(keys(index))
        xValues(index) = CDbl(point(0))
        yValues(index) = CDbl(point(1))
        zValues(index) = CDbl(point(2))
    Next index

    SortAnchors nodes, xValues, yValues, zValues, anchorCount

    For index = 1 To anchorCount - 1
        If nodes(index) <= nodes(index - 1) Then
            Err.Raise vbObjectError + 2002, , _
                "Mesh node numbers must be unique and increasing after sorting."
        End If
    Next index
End Sub

Private Function FindAnchorBlockLastRow(ByVal ws As Worksheet) As Long
    Dim rowNumber As Long
    Dim foundData As Boolean
    Dim blankRun As Long

    For rowNumber = 1 To CURVE_FIRST_ROW - 1
        If HasNumericValue(ws.Cells(rowNumber, 1)) Or _
           HasNumericValue(ws.Cells(rowNumber, 7)) Then
            foundData = True
            blankRun = 0
            FindAnchorBlockLastRow = rowNumber
        ElseIf foundData Then
            blankRun = blankRun + 1
            If blankRun >= 2 Then Exit For
        End If
    Next rowNumber
End Function

Private Function IsCompleteAnchorRow(ByVal ws As Worksheet, _
                                     ByVal rowNumber As Long, _
                                     ByVal baseColumn As Long) As Boolean
    Dim offset As Long
    For offset = 0 To 3
        If Not HasNumericValue(ws.Cells(rowNumber, baseColumn + offset)) Then Exit Function
    Next offset
    IsCompleteAnchorRow = True
End Function

Private Function HasNumericValue(ByVal cell As Range) As Boolean
    If IsError(cell.Value2) Then Exit Function
    If Len(cell.Value2) = 0 Then Exit Function
    HasNumericValue = IsNumeric(cell.Value2)
End Function

Private Sub SortAnchors(ByRef nodes() As Long, _
                        ByRef xValues() As Double, _
                        ByRef yValues() As Double, _
                        ByRef zValues() As Double, _
                        ByVal itemCount As Long)
    Dim i As Long
    Dim j As Long
    Dim nodeTemp As Long
    Dim valueTemp As Double

    For i = 1 To itemCount - 1
        j = i
        Do While j > 0 And nodes(j) < nodes(j - 1)
            nodeTemp = nodes(j - 1)
            nodes(j - 1) = nodes(j)
            nodes(j) = nodeTemp

            valueTemp = xValues(j - 1)
            xValues(j - 1) = xValues(j)
            xValues(j) = valueTemp

            valueTemp = yValues(j - 1)
            yValues(j - 1) = yValues(j)
            yValues(j) = valueTemp

            valueTemp = zValues(j - 1)
            zValues(j - 1) = zValues(j)
            zValues(j) = valueTemp

            j = j - 1
        Loop
    Next i
End Sub

Private Function LoadCurves(ByVal ws As Worksheet, _
                            ByRef anchorNodes() As Long, _
                            ByVal anchorCount As Long) As Collection
    Dim result As New Collection
    Dim rowNumber As Long
    Dim startNode As Long
    Dim endNode As Long
    Dim direction As Long
    Dim previousEnd As Long

    rowNumber = CURVE_FIRST_ROW
    Do While rowNumber <= ws.Rows.Count
        If IsCurveRowBlank(ws, rowNumber) Then Exit Do

        If Not HasNumericValue(ws.Cells(rowNumber, 1)) Or _
           Not HasNumericValue(ws.Cells(rowNumber, 2)) Or _
           Not HasNumericValue(ws.Cells(rowNumber, 4)) Or _
           Not HasNumericValue(ws.Cells(rowNumber, 5)) Or _
           Not HasNumericValue(ws.Cells(rowNumber, 9)) Then
            Err.Raise vbObjectError + 2003, , _
                "Curve row " & CStr(rowNumber) & _
                " must contain start node, end node, center X, center Y, and type."
        End If

        startNode = CLng(ws.Cells(rowNumber, 1).Value2)
        endNode = CLng(ws.Cells(rowNumber, 2).Value2)
        direction = CLng(ws.Cells(rowNumber, 9).Value2)

        If endNode <= startNode Then
            Err.Raise vbObjectError + 2004, , _
                "Curve row " & CStr(rowNumber) & " has an invalid node interval."
        End If
        If direction <> -1 And direction <> 1 Then
            Err.Raise vbObjectError + 2005, , _
                "Curve type in I" & CStr(rowNumber) & " must be -1 or 1."
        End If
        If FindExactAnchorIndex(anchorNodes, anchorCount, startNode) < 0 Or _
           FindExactAnchorIndex(anchorNodes, anchorCount, endNode) < 0 Then
            Err.Raise vbObjectError + 2006, , _
                "Curve row " & CStr(rowNumber) & _
                " refers to a start or end node that is not in the anchor table."
        End If
        If result.Count > 0 And startNode <= previousEnd Then
            Err.Raise vbObjectError + 2007, , _
                "Curve definitions overlap or are not ordered at row " & CStr(rowNumber) & "."
        End If

        result.Add Array( _
            startNode, _
            endNode, _
            CDbl(ws.Cells(rowNumber, 4).Value2), _
            CDbl(ws.Cells(rowNumber, 5).Value2), _
            direction)
        previousEnd = endNode
        rowNumber = rowNumber + 1
    Loop

    Set LoadCurves = result
End Function

Private Function IsCurveRowBlank(ByVal ws As Worksheet, ByVal rowNumber As Long) As Boolean
    IsCurveRowBlank = _
        Len(ws.Cells(rowNumber, 1).Value2) = 0 And _
        Len(ws.Cells(rowNumber, 2).Value2) = 0 And _
        Len(ws.Cells(rowNumber, 4).Value2) = 0 And _
        Len(ws.Cells(rowNumber, 5).Value2) = 0 And _
        Len(ws.Cells(rowNumber, 9).Value2) = 0
End Function

Private Sub BuildOutputArray(ByVal ws As Worksheet, _
                             ByRef anchorNodes() As Long, _
                             ByRef anchorX() As Double, _
                             ByRef anchorY() As Double, _
                             ByRef anchorZ() As Double, _
                             ByVal anchorCount As Long, _
                             ByVal curves As Collection, _
                             ByRef outputValues() As Variant, _
                             ByRef outputStartRow As Long, _
                             ByRef outputLastRow As Long, _
                             ByRef generatedCount As Long)
    Dim matrixRow As Long
    Dim nodeNumber As Long
    Dim xValue As Double
    Dim yValue As Double
    Dim zValue As Double
    Dim outputCount As Long

    ' Z:AC is sized directly from the current mesh node range. Column T is
    ' intentionally not used, so a larger or smaller imported mesh adjusts
    ' the output without requiring any manual fill-down work.
    outputStartRow = DEFAULT_OUTPUT_FIRST_ROW
    outputCount = anchorNodes(anchorCount - 1) - anchorNodes(0) + 1
    If outputCount < 1 Or outputStartRow + outputCount - 1 > ws.Rows.Count Then
        Err.Raise vbObjectError + 2008, , "The generated node range does not fit on the worksheet."
    End If
    outputLastRow = outputStartRow + outputCount - 1
    ReDim outputValues(1 To outputCount, 1 To 4)

    For matrixRow = 1 To outputCount
        nodeNumber = anchorNodes(0) + matrixRow - 1
        ComputeCoordinate nodeNumber, anchorNodes, anchorX, anchorY, anchorZ, _
                          anchorCount, curves, xValue, yValue, zValue
        outputValues(matrixRow, 1) = nodeNumber
        outputValues(matrixRow, 2) = xValue
        outputValues(matrixRow, 3) = yValue
        outputValues(matrixRow, 4) = zValue
    Next matrixRow
    generatedCount = outputCount
End Sub

Private Function FindLastOutputRow(ByVal ws As Worksheet) As Long
    Dim columnNumber As Long
    Dim candidateRow As Long

    For columnNumber = OUTPUT_FIRST_COL To OUTPUT_FIRST_COL + 3
        candidateRow = ws.Cells(ws.Rows.Count, columnNumber).End(xlUp).Row
        If candidateRow > FindLastOutputRow Then FindLastOutputRow = candidateRow
    Next columnNumber

    If FindLastOutputRow < DEFAULT_OUTPUT_FIRST_ROW Then
        FindLastOutputRow = DEFAULT_OUTPUT_FIRST_ROW - 1
    End If
End Function

Private Sub ExtendOutputFormatting(ByVal ws As Worksheet, _
                                   ByVal outputStartRow As Long, _
                                   ByVal outputLastRow As Long, _
                                   ByVal previousOutputLastRow As Long)
    If outputLastRow <= previousOutputLastRow Then Exit Sub
    If previousOutputLastRow < outputStartRow Then Exit Sub

    ws.Range(ws.Cells(previousOutputLastRow, OUTPUT_FIRST_COL), _
             ws.Cells(previousOutputLastRow, OUTPUT_FIRST_COL + 3)).Copy
    ws.Range(ws.Cells(previousOutputLastRow + 1, OUTPUT_FIRST_COL), _
             ws.Cells(outputLastRow, OUTPUT_FIRST_COL + 3)).PasteSpecial Paste:=xlPasteFormats
    Application.CutCopyMode = False
End Sub

Private Sub UpdateSpringSheet(ByVal sourceSheet As Worksheet, _
                              ByVal sourceFirstRow As Long, _
                              ByVal sourceLastRow As Long, _
                              ByVal nodeCount As Long)
    Dim springSheet As Worksheet
    Dim springLastRow As Long
    Dim springInteriorLastRow As Long
    Dim previousSpringLastRow As Long
    Dim clearLastRow As Long
    Dim seedFormulaOrValue(7 To 25) As Variant
    Dim seedValues(7 To 25) As Variant
    Dim seedItem As Variant
    Dim seedValue As Variant
    Dim springColumn As Long
    Dim formulaLastRow As Long
    Dim markerF As Variant

    Set springSheet = ThisWorkbook.Worksheets(SPRING_SHEET)
    springLastRow = SPRING_OUTPUT_FIRST_ROW + nodeCount - 1
    If springLastRow > springSheet.Rows.Count Then
        Err.Raise vbObjectError + 2011, , _
            "The Spring2-10 output does not fit on the worksheet."
    End If

    ' Capture every G:Y seed independently before clearing the old dynamic
    ' area. FormulaR1C1 preserves each column's own relative-reference pattern.
    For springColumn = 7 To 25
        seedFormulaOrValue(springColumn) = _
            springSheet.Cells(6, springColumn).FormulaR1C1
        seedValues(springColumn) = springSheet.Cells(6, springColumn).Value2
    Next springColumn

    If Left$(CStr(seedFormulaOrValue(7)), 1) <> "=" Or _
       Left$(CStr(seedFormulaOrValue(25)), 1) <> "=" Then
        Err.Raise vbObjectError + 2012, , _
            "Spring2-10 formulas in G6:Y6 are missing or invalid."
    End If

    markerF = springSheet.Range("F5").Value2

    previousSpringLastRow = FindLastSpringOutputRow(springSheet)
    ExtendSpringFormatting springSheet, springLastRow, previousSpringLastRow

    clearLastRow = previousSpringLastRow
    If springLastRow > clearLastRow Then clearLastRow = springLastRow
    If clearLastRow >= SPRING_OUTPUT_FIRST_ROW Then
        springSheet.Range("B" & CStr(SPRING_OUTPUT_FIRST_ROW) & _
                          ":Y" & CStr(clearLastRow)).ClearContents
    End If

    springSheet.Range("B" & CStr(SPRING_OUTPUT_FIRST_ROW) & _
                      ":E" & CStr(springLastRow)).Value2 = _
        sourceSheet.Range(sourceSheet.Cells(sourceFirstRow, OUTPUT_FIRST_COL), _
                          sourceSheet.Cells(sourceLastRow, OUTPUT_FIRST_COL + 3)).Value2

    FillSpringMarker springSheet, "F", markerF, springLastRow

    springInteriorLastRow = springLastRow - 1
    For springColumn = 7 To 25  ' G:Y
        seedItem = seedFormulaOrValue(springColumn)
        If VarType(seedItem) = vbString And Left$(CStr(seedItem), 1) = "=" Then
            ' G:I describe the segment ending at the current node and include
            ' the final node. Later calculated columns need a next-row vector.
            If springColumn <= 9 Then
                formulaLastRow = springLastRow
            Else
                formulaLastRow = springInteriorLastRow
            End If

            If formulaLastRow >= 6 Then
                springSheet.Range(springSheet.Cells(6, springColumn), _
                                  springSheet.Cells(formulaLastRow, springColumn)).FormulaR1C1 = _
                    CStr(seedItem)
            End If
        Else
            seedValue = seedValues(springColumn)
            If Not IsError(seedValue) Then
                If Not IsNull(seedValue) Then
                    If Len(CStr(seedValue)) > 0 Then
                        springSheet.Range(springSheet.Cells(SPRING_OUTPUT_FIRST_ROW, springColumn), _
                                          springSheet.Cells(springLastRow, springColumn)).Value2 = seedValue
                    End If
                End If
            End If
        End If
    Next springColumn

    springSheet.Range("B" & CStr(SPRING_OUTPUT_FIRST_ROW) & _
                      ":Y" & CStr(springLastRow)).Calculate

    ExportSpringTextFile springSheet, springInteriorLastRow
End Sub

Private Sub ExportSpringTextFile(ByVal springSheet As Worksheet, _
                                 ByVal springInteriorLastRow As Long)
    Dim rawValues As Variant
    Dim blocks() As String
    Dim blockCount As Long
    Dim blockIndex As Long
    Dim cellValue As Variant
    Dim generatedText As String
    Dim existingText As String
    Dim footerText As String
    Dim firstSpringNode As Long
    Dim lastSpringNode As Long
    Dim outputPath As String

    If springInteriorLastRow < 6 Then
        Err.Raise vbObjectError + 2013, , _
            "Spring2-10 does not contain any interior nodes to export."
    End If

    blockCount = springInteriorLastRow - 5
    rawValues = springSheet.Range("Y6:Y" & CStr(springInteriorLastRow)).Value2
    ReDim blocks(0 To blockCount - 1)

    For blockIndex = 1 To blockCount
        If blockCount = 1 Then
            cellValue = rawValues
        Else
            cellValue = rawValues(blockIndex, 1)
        End If

        If IsError(cellValue) Or Len(CStr(cellValue)) = 0 Then
            Err.Raise vbObjectError + 2014, , _
                "Spring2-10 column Y is blank or invalid at row " & _
                CStr(blockIndex + 5) & "."
        End If

        blocks(blockIndex - 1) = Replace(CStr(cellValue), "hzw", vbCrLf, _
                                         1, -1, vbTextCompare)
    Next blockIndex

    generatedText = Join(blocks, vbCrLf)
    outputPath = WorkbookFolderFilePath(SPRING_TEXT_FILE_NAME)
    existingText = ReadEntireTextFile(outputPath)
    footerText = ExtractSpringFooter(existingText)

    firstSpringNode = CLng(springSheet.Cells(6, 2).Value2)
    lastSpringNode = CLng(springSheet.Cells(springInteriorLastRow, 2).Value2)
    footerText = UpdateSpringFooterRange(footerText, _
                                         firstSpringNode + 600000, _
                                         lastSpringNode + 600000)

    WriteEntireTextFileSafely outputPath, _
                              generatedText & vbCrLf & footerText & vbCrLf
End Sub

Private Sub UpdateSpringDistanceSheet(ByVal sourceSheet As Worksheet, _
                                      ByVal sourceFirstRow As Long, _
                                      ByVal sourceLastRow As Long, _
                                      ByVal nodeCount As Long)
    Dim distanceSheet As Worksheet
    Dim distanceLastRow As Long
    Dim previousDistanceLastRow As Long
    Dim clearLastRow As Long
    Dim seedFValue As Variant
    Dim formulaF As String
    Dim formulaK As String
    Dim seedFormulaOrValue(14 To 25) As Variant
    Dim seedValues(14 To 25) As Variant
    Dim seedItem As Variant
    Dim seedValue As Variant
    Dim distanceColumn As Long

    Set distanceSheet = ThisWorkbook.Worksheets(SPRING_DISTANCE_SHEET)
    distanceLastRow = SPRING_DISTANCE_FIRST_ROW + nodeCount - 1
    If distanceLastRow > distanceSheet.Rows.Count Then
        Err.Raise vbObjectError + 2016, , _
            "The SpringDistance2-10 output does not fit on the worksheet."
    End If

    ' Preserve each worksheet seed before clearing the prior dynamic area.
    ' F9 is the cumulative-distance start; F10 is its repeating formula.
    seedFValue = distanceSheet.Range("F9").Value2
    formulaF = CStr(distanceSheet.Range("F10").FormulaR1C1)
    formulaK = CStr(distanceSheet.Range("K9").FormulaR1C1)
    For distanceColumn = 14 To 25  ' N:Y
        seedFormulaOrValue(distanceColumn) = _
            distanceSheet.Cells(9, distanceColumn).FormulaR1C1
        seedValues(distanceColumn) = _
            distanceSheet.Cells(9, distanceColumn).Value2
    Next distanceColumn

    If Left$(formulaF, 1) <> "=" Or Left$(formulaK, 1) <> "=" Then
        Err.Raise vbObjectError + 2017, , _
            "SpringDistance2-10 seed formulas in F10 or K9 are missing."
    End If

    previousDistanceLastRow = FindLastSpringDistanceRow(distanceSheet)
    ExtendSpringDistanceFormatting distanceSheet, distanceLastRow, _
                                  previousDistanceLastRow

    clearLastRow = previousDistanceLastRow
    If distanceLastRow > clearLastRow Then clearLastRow = distanceLastRow
    If clearLastRow >= SPRING_DISTANCE_FIRST_ROW Then
        distanceSheet.Range("F" & CStr(SPRING_DISTANCE_FIRST_ROW) & _
                            ":K" & CStr(clearLastRow)).ClearContents
        distanceSheet.Range("N" & CStr(SPRING_DISTANCE_FIRST_ROW) & _
                            ":Y" & CStr(clearLastRow)).ClearContents
    End If

    distanceSheet.Range("G" & CStr(SPRING_DISTANCE_FIRST_ROW) & _
                        ":J" & CStr(distanceLastRow)).Value2 = _
        sourceSheet.Range(sourceSheet.Cells(sourceFirstRow, OUTPUT_FIRST_COL), _
                          sourceSheet.Cells(sourceLastRow, OUTPUT_FIRST_COL + 3)).Value2

    distanceSheet.Range("F9").Value2 = seedFValue
    If nodeCount >= 2 Then
        distanceSheet.Range("F10:F" & CStr(distanceLastRow)).FormulaR1C1 = formulaF
    End If
    distanceSheet.Range("K9:K" & CStr(distanceLastRow)).FormulaR1C1 = formulaK

    For distanceColumn = 14 To 25  ' N:Y
        seedItem = seedFormulaOrValue(distanceColumn)
        If VarType(seedItem) = vbString And Left$(CStr(seedItem), 1) = "=" Then
            distanceSheet.Range( _
                distanceSheet.Cells(SPRING_DISTANCE_FIRST_ROW, distanceColumn), _
                distanceSheet.Cells(distanceLastRow, distanceColumn)).FormulaR1C1 = _
                    CStr(seedItem)
        Else
            seedValue = seedValues(distanceColumn)
            If Not IsError(seedValue) Then
                If Not IsNull(seedValue) Then
                    If Len(CStr(seedValue)) > 0 Then
                        distanceSheet.Range( _
                            distanceSheet.Cells(SPRING_DISTANCE_FIRST_ROW, distanceColumn), _
                            distanceSheet.Cells(distanceLastRow, distanceColumn)).Value2 = seedValue
                    End If
                End If
            End If
        End If
    Next distanceColumn

    distanceSheet.Range("F" & CStr(SPRING_DISTANCE_FIRST_ROW) & _
                        ":K" & CStr(distanceLastRow)).Calculate
    distanceSheet.Range("N" & CStr(SPRING_DISTANCE_FIRST_ROW) & _
                        ":Y" & CStr(distanceLastRow)).Calculate

    ExportInitialDisplacementTextFile distanceSheet, distanceLastRow
End Sub

Private Sub ExportInitialDisplacementTextFile( _
            ByVal distanceSheet As Worksheet, _
            ByVal distanceLastRow As Long)
    Dim rawNodes As Variant
    Dim rawSeparators As Variant
    Dim rawDisplacements As Variant
    Dim outputLines() As String
    Dim rowCount As Long
    Dim itemIndex As Long
    Dim nodeValue As Variant
    Dim separatorValue As Variant
    Dim displacementValue As Variant
    Dim displacementText As String
    Dim decimalSeparator As String
    Dim headerText As String
    Dim outputPath As String

    If distanceLastRow < SPRING_DISTANCE_FIRST_ROW Then
        Err.Raise vbObjectError + 2018, , _
            "SpringDistance2-10 does not contain any rows to export."
    End If

    rowCount = distanceLastRow - SPRING_DISTANCE_FIRST_ROW + 1
    rawNodes = distanceSheet.Range("N9:N" & CStr(distanceLastRow)).Value2
    rawSeparators = distanceSheet.Range("O9:O" & CStr(distanceLastRow)).Value2
    rawDisplacements = distanceSheet.Range("P9:P" & CStr(distanceLastRow)).Value2
    ReDim outputLines(0 To rowCount)

    headerText = CStr(distanceSheet.Range("N8").Value2)
    If Len(headerText) = 0 Then
        Err.Raise vbObjectError + 2019, , _
            "SpringDistance2-10 header N8 is blank."
    End If
    outputLines(0) = headerText & vbTab & vbTab

    decimalSeparator = Application.International(xlDecimalSeparator)

    For itemIndex = 1 To rowCount
        If rowCount = 1 Then
            nodeValue = rawNodes
            separatorValue = rawSeparators
            displacementValue = rawDisplacements
        Else
            nodeValue = rawNodes(itemIndex, 1)
            separatorValue = rawSeparators(itemIndex, 1)
            displacementValue = rawDisplacements(itemIndex, 1)
        End If

        If IsError(nodeValue) Then
            Err.Raise vbObjectError + 2020, , _
                "SpringDistance2-10 column N is blank or invalid at row " & _
                CStr(itemIndex + SPRING_DISTANCE_FIRST_ROW - 1) & "."
        End If
        If Not IsNumeric(nodeValue) Then
            Err.Raise vbObjectError + 2020, , _
                "SpringDistance2-10 column N is blank or invalid at row " & _
                CStr(itemIndex + SPRING_DISTANCE_FIRST_ROW - 1) & "."
        End If
        If IsError(separatorValue) Then
            Err.Raise vbObjectError + 2021, , _
                "SpringDistance2-10 column O is blank or invalid at row " & _
                CStr(itemIndex + SPRING_DISTANCE_FIRST_ROW - 1) & "."
        End If
        If Len(CStr(separatorValue)) = 0 Then
            Err.Raise vbObjectError + 2021, , _
                "SpringDistance2-10 column O is blank or invalid at row " & _
                CStr(itemIndex + SPRING_DISTANCE_FIRST_ROW - 1) & "."
        End If
        If IsError(displacementValue) Then
            Err.Raise vbObjectError + 2022, , _
                "SpringDistance2-10 column P is blank or invalid at row " & _
                CStr(itemIndex + SPRING_DISTANCE_FIRST_ROW - 1) & "."
        End If
        If Not IsNumeric(displacementValue) Then
            Err.Raise vbObjectError + 2022, , _
                "SpringDistance2-10 column P is blank or invalid at row " & _
                CStr(itemIndex + SPRING_DISTANCE_FIRST_ROW - 1) & "."
        End If

        displacementText = Format$(CDbl(displacementValue), "0.000")
        If decimalSeparator <> "." Then
            displacementText = Replace(displacementText, decimalSeparator, ".")
        End If

        outputLines(itemIndex) = _
            CStr(CLng(nodeValue)) & vbTab & _
            CStr(separatorValue) & vbTab & displacementText
    Next itemIndex

    outputPath = WorkbookFolderFilePath( _
        INITIAL_DISPLACEMENT_TEXT_FILE_NAME)
    WriteEntireTextFileSafely outputPath, _
                              Join(outputLines, vbCrLf)
End Sub

Private Sub UpdateSheet10(ByVal sourceSheet As Worksheet, _
                          ByVal sourceFirstRow As Long, _
                          ByVal sourceLastRow As Long, _
                          ByVal nodeCount As Long)
    Dim outputSheet As Worksheet
    Dim sheet10LastRow As Long
    Dim previousSheet10LastRow As Long
    Dim clearLastRow As Long
    Dim seedBValue As Variant
    Dim seedNValue As Variant
    Dim headerMValue As Variant
    Dim formulaA As String
    Dim formulaB As String
    Dim formulaG As String
    Dim formulaH As String
    Dim formulaI As String
    Dim formulaJFirst As String
    Dim formulaJRepeat As String
    Dim formulaK As String
    Dim formulaL As String
    Dim formulaM As String
    Dim formulaO As String

    Set outputSheet = ThisWorkbook.Worksheets(SHEET10_SHEET)
    sheet10LastRow = SHEET10_FIRST_ROW + nodeCount - 1
    If sheet10LastRow > outputSheet.Rows.Count Then
        Err.Raise vbObjectError + 2023, , _
            "The Sheet10 output does not fit on the worksheet."
    End If

    ' Row 9 is the boundary/start row. Most repeating formulas begin at row
    ' 10, and J11 is the repeating J-column seed because J10 is a special
    ' first-data-row formula in the existing worksheet.
    formulaA = CStr(outputSheet.Range("A9").FormulaR1C1)
    seedBValue = outputSheet.Range("B9").Value2
    formulaB = CStr(outputSheet.Range("B10").FormulaR1C1)
    formulaG = CStr(outputSheet.Range("G10").FormulaR1C1)
    formulaH = CStr(outputSheet.Range("H10").FormulaR1C1)
    formulaI = CStr(outputSheet.Range("I10").FormulaR1C1)
    formulaJFirst = CStr(outputSheet.Range("J10").FormulaR1C1)
    formulaJRepeat = CStr(outputSheet.Range("J11").FormulaR1C1)
    formulaK = CStr(outputSheet.Range("K10").FormulaR1C1)
    formulaL = CStr(outputSheet.Range("L10").FormulaR1C1)
    formulaM = CStr(outputSheet.Range("M10").FormulaR1C1)
    seedNValue = outputSheet.Range("N10").Value2
    formulaO = CStr(outputSheet.Range("O10").FormulaR1C1)
    headerMValue = outputSheet.Range("M9").Value2

    If Left$(formulaA, 1) <> "=" Or _
       Left$(formulaB, 1) <> "=" Or _
       Left$(formulaG, 1) <> "=" Or _
       Left$(formulaH, 1) <> "=" Or _
       Left$(formulaI, 1) <> "=" Or _
       Left$(formulaJFirst, 1) <> "=" Or _
       Left$(formulaJRepeat, 1) <> "=" Or _
       Left$(formulaK, 1) <> "=" Or _
       Left$(formulaL, 1) <> "=" Or _
       Left$(formulaM, 1) <> "=" Or _
       Left$(formulaO, 1) <> "=" Then
        Err.Raise vbObjectError + 2024, , _
            "One or more Sheet10 seed formulas are missing or invalid."
    End If

    previousSheet10LastRow = FindLastSheet10Row(outputSheet)
    ExtendSheet10Formatting outputSheet, sheet10LastRow, _
                             previousSheet10LastRow

    clearLastRow = previousSheet10LastRow
    If sheet10LastRow > clearLastRow Then clearLastRow = sheet10LastRow
    If clearLastRow >= SHEET10_FIRST_ROW Then
        outputSheet.Range("A" & CStr(SHEET10_FIRST_ROW) & _
                          ":O" & CStr(clearLastRow)).ClearContents
    End If

    outputSheet.Range("C" & CStr(SHEET10_FIRST_ROW) & _
                      ":F" & CStr(sheet10LastRow)).Value2 = _
        sourceSheet.Range(sourceSheet.Cells(sourceFirstRow, OUTPUT_FIRST_COL), _
                          sourceSheet.Cells(sourceLastRow, OUTPUT_FIRST_COL + 3)).Value2

    outputSheet.Range("A9:A" & CStr(sheet10LastRow)).FormulaR1C1 = formulaA
    outputSheet.Range("B9").Value2 = seedBValue
    outputSheet.Range("M9").Value2 = headerMValue

    If sheet10LastRow >= 10 Then
        outputSheet.Range("B10:B" & CStr(sheet10LastRow)).FormulaR1C1 = formulaB
        outputSheet.Range("G10:G" & CStr(sheet10LastRow)).FormulaR1C1 = formulaG
        outputSheet.Range("H10:H" & CStr(sheet10LastRow)).FormulaR1C1 = formulaH
        outputSheet.Range("I10:I" & CStr(sheet10LastRow)).FormulaR1C1 = formulaI
        outputSheet.Range("J10").FormulaR1C1 = formulaJFirst
        outputSheet.Range("K10:K" & CStr(sheet10LastRow)).FormulaR1C1 = formulaK
        outputSheet.Range("L10:L" & CStr(sheet10LastRow)).FormulaR1C1 = formulaL
        outputSheet.Range("M10:M" & CStr(sheet10LastRow)).FormulaR1C1 = formulaM
        outputSheet.Range("N10:N" & CStr(sheet10LastRow)).Value2 = seedNValue
        outputSheet.Range("O10:O" & CStr(sheet10LastRow)).FormulaR1C1 = formulaO
    End If

    If sheet10LastRow >= 11 Then
        outputSheet.Range("J11:J" & CStr(sheet10LastRow)).FormulaR1C1 = _
            formulaJRepeat
    End If

    ' Column A must be calculated before H5 is pointed to its dynamic last row.
    outputSheet.Range("A9:B" & CStr(sheet10LastRow)).Calculate
    outputSheet.Range("H5").Formula = "=A" & CStr(sheet10LastRow)
    outputSheet.Range("H5:I5").Calculate
    outputSheet.Range("G9:O" & CStr(sheet10LastRow)).Calculate

    ExportTensionAdjustmentTextFile outputSheet, sheet10LastRow
End Sub

Private Function FindLastSheet10Row(ByVal outputSheet As Worksheet) As Long
    Dim columnNumber As Long
    Dim candidateRow As Long

    For columnNumber = 1 To 15  ' A:O dynamic table only
        candidateRow = outputSheet.Cells(outputSheet.Rows.Count, _
                                          columnNumber).End(xlUp).Row
        If candidateRow > FindLastSheet10Row Then
            FindLastSheet10Row = candidateRow
        End If
    Next columnNumber

    If FindLastSheet10Row < SHEET10_FIRST_ROW Then
        FindLastSheet10Row = SHEET10_FIRST_ROW - 1
    End If
End Function

Private Sub ExtendSheet10Formatting(ByVal outputSheet As Worksheet, _
                                    ByVal sheet10LastRow As Long, _
                                    ByVal previousSheet10LastRow As Long)
    Dim formattingFirstRow As Long

    If sheet10LastRow <= previousSheet10LastRow Then Exit Sub

    formattingFirstRow = previousSheet10LastRow + 1
    If formattingFirstRow < 11 Then formattingFirstRow = 11
    If formattingFirstRow > sheet10LastRow Then Exit Sub

    outputSheet.Range("A11:O11").Copy
    outputSheet.Range("A" & CStr(formattingFirstRow) & _
                      ":O" & CStr(sheet10LastRow)).PasteSpecial Paste:=xlPasteFormats
    Application.CutCopyMode = False
End Sub

Private Sub ExportTensionAdjustmentTextFile( _
            ByVal outputSheet As Worksheet, _
            ByVal sheet10LastRow As Long)
    Dim rawNodes As Variant
    Dim rawSeparators As Variant
    Dim rawLoads As Variant
    Dim outputLines() As String
    Dim dataRowCount As Long
    Dim itemIndex As Long
    Dim nodeValue As Variant
    Dim separatorValue As Variant
    Dim loadValue As Variant
    Dim loadText As String
    Dim decimalSeparator As String
    Dim headerText As String
    Dim existingText As String
    Dim specialPrefix As String
    Dim specialCellValue As Variant
    Dim specialValueText As String
    Dim outputPath As String

    headerText = NormalizeLineBreaks(CStr(outputSheet.Range("M9").Value2))
    Do While Len(headerText) >= 2 And Right$(headerText, 2) = vbCrLf
        headerText = Left$(headerText, Len(headerText) - 2)
    Loop
    If Len(headerText) = 0 Then
        Err.Raise vbObjectError + 2025, , _
            "Sheet10 tension-adjustment header M9 is blank."
    End If

    specialCellValue = outputSheet.Range("I3").Value2
    If IsError(specialCellValue) Then
        Err.Raise vbObjectError + 2026, , _
            "Sheet10 special tension value I3 is invalid."
    End If
    If Not IsNumeric(specialCellValue) Then
        Err.Raise vbObjectError + 2026, , _
            "Sheet10 special tension value I3 is blank or invalid."
    End If

    outputPath = WorkbookFolderFilePath(TENSION_ADJUSTMENT_TEXT_FILE_NAME)
    existingText = ReadEntireTextFile(outputPath)
    specialPrefix = ExtractTensionSpecialPrefix(existingText)
    specialValueText = CStr(CLng(Application.WorksheetFunction.Round( _
        -Abs(CDbl(specialCellValue)), 0)))

    If sheet10LastRow >= 10 Then
        dataRowCount = sheet10LastRow - 9
        rawNodes = outputSheet.Range("M10:M" & CStr(sheet10LastRow)).Value2
        rawSeparators = outputSheet.Range("N10:N" & CStr(sheet10LastRow)).Value2
        rawLoads = outputSheet.Range("O10:O" & CStr(sheet10LastRow)).Value2
    Else
        dataRowCount = 0
    End If

    ReDim outputLines(0 To dataRowCount + 1)
    outputLines(0) = headerText
    outputLines(1) = specialPrefix & specialValueText
    decimalSeparator = Application.International(xlDecimalSeparator)

    For itemIndex = 1 To dataRowCount
        If dataRowCount = 1 Then
            nodeValue = rawNodes
            separatorValue = rawSeparators
            loadValue = rawLoads
        Else
            nodeValue = rawNodes(itemIndex, 1)
            separatorValue = rawSeparators(itemIndex, 1)
            loadValue = rawLoads(itemIndex, 1)
        End If

        If IsError(nodeValue) Then
            Err.Raise vbObjectError + 2027, , _
                "Sheet10 column M is invalid at row " & CStr(itemIndex + 9) & "."
        End If
        If Not IsNumeric(nodeValue) Then
            Err.Raise vbObjectError + 2027, , _
                "Sheet10 column M is blank or invalid at row " & _
                CStr(itemIndex + 9) & "."
        End If
        If IsError(separatorValue) Then
            Err.Raise vbObjectError + 2028, , _
                "Sheet10 column N is invalid at row " & CStr(itemIndex + 9) & "."
        End If
        If Len(CStr(separatorValue)) = 0 Then
            Err.Raise vbObjectError + 2028, , _
                "Sheet10 column N is blank at row " & CStr(itemIndex + 9) & "."
        End If
        If IsError(loadValue) Then
            Err.Raise vbObjectError + 2029, , _
                "Sheet10 column O is invalid at row " & CStr(itemIndex + 9) & "."
        End If
        If Not IsNumeric(loadValue) Then
            Err.Raise vbObjectError + 2029, , _
                "Sheet10 column O is blank or invalid at row " & _
                CStr(itemIndex + 9) & "."
        End If

        loadText = Format$(CDbl(loadValue), "0.000")
        If decimalSeparator <> "." Then
            loadText = Replace(loadText, decimalSeparator, ".")
        End If

        outputLines(itemIndex + 1) = _
            CStr(CLng(nodeValue)) & vbTab & _
            CStr(separatorValue) & vbTab & loadText
    Next itemIndex

    WriteEntireTextFileSafely outputPath, _
                              Join(outputLines, vbCrLf)
End Sub

Private Function WorkbookFolderFilePath(ByVal fileName As String) As String
    Dim workbookFolder As String

    workbookFolder = ThisWorkbook.Path
    If Len(workbookFolder) = 0 Then
        Err.Raise vbObjectError + 2031, , _
            "Save this workbook before running a workflow that exports text files."
    End If

    If Right$(workbookFolder, 1) = Application.PathSeparator Then
        WorkbookFolderFilePath = workbookFolder & fileName
    Else
        WorkbookFolderFilePath = workbookFolder & _
                                 Application.PathSeparator & fileName
    End If
End Function

Private Function ExtractTensionSpecialPrefix( _
                 ByVal existingText As String) As String
    Dim lines As Variant
    Dim lineIndex As Long
    Dim lineValue As String
    Dim firstComma As Long
    Dim secondComma As Long
    Dim remainder As String
    Dim whitespaceLength As Long

    If Len(existingText) > 0 Then
        lines = Split(NormalizeLineBreaks(existingText), vbCrLf)
        For lineIndex = LBound(lines) To UBound(lines)
            lineValue = CStr(lines(lineIndex))
            firstComma = InStr(1, lineValue, ",", vbBinaryCompare)
            If firstComma > 0 Then
                secondComma = InStr(firstComma + 1, lineValue, ",", vbBinaryCompare)
                If secondComma > firstComma Then
                    If Trim$(Left$(lineValue, firstComma - 1)) = "1" And _
                       Trim$(Mid$(lineValue, firstComma + 1, _
                                 secondComma - firstComma - 1)) = "1" Then
                        remainder = Mid$(lineValue, secondComma + 1)
                        whitespaceLength = Len(remainder) - Len(LTrim$(remainder))
                        ExtractTensionSpecialPrefix = _
                            Left$(lineValue, secondComma) & _
                            Left$(remainder, whitespaceLength)
                        If Len(ExtractTensionSpecialPrefix) > 0 Then Exit Function
                    End If
                End If
            End If
        Next lineIndex
    End If

    ExtractTensionSpecialPrefix = "1,        1,   "
End Function

Private Function FindLastSpringDistanceRow(ByVal distanceSheet As Worksheet) As Long
    Dim distanceColumn As Long
    Dim candidateRow As Long

    For distanceColumn = 6 To 11  ' F:K
        candidateRow = distanceSheet.Cells(distanceSheet.Rows.Count, _
                                            distanceColumn).End(xlUp).Row
        If candidateRow > FindLastSpringDistanceRow Then
            FindLastSpringDistanceRow = candidateRow
        End If
    Next distanceColumn

    For distanceColumn = 14 To 25  ' N:Y
        candidateRow = distanceSheet.Cells(distanceSheet.Rows.Count, _
                                            distanceColumn).End(xlUp).Row
        If candidateRow > FindLastSpringDistanceRow Then
            FindLastSpringDistanceRow = candidateRow
        End If
    Next distanceColumn

    If FindLastSpringDistanceRow < SPRING_DISTANCE_FIRST_ROW Then
        FindLastSpringDistanceRow = SPRING_DISTANCE_FIRST_ROW - 1
    End If
End Function

Private Sub ExtendSpringDistanceFormatting(ByVal distanceSheet As Worksheet, _
                                           ByVal distanceLastRow As Long, _
                                           ByVal previousDistanceLastRow As Long)
    If distanceLastRow <= previousDistanceLastRow Then Exit Sub

    distanceSheet.Range("F9:K9").Copy
    distanceSheet.Range("F" & CStr(previousDistanceLastRow + 1) & _
                        ":K" & CStr(distanceLastRow)).PasteSpecial Paste:=xlPasteFormats

    distanceSheet.Range("N9:Y9").Copy
    distanceSheet.Range("N" & CStr(previousDistanceLastRow + 1) & _
                        ":Y" & CStr(distanceLastRow)).PasteSpecial Paste:=xlPasteFormats
    Application.CutCopyMode = False
End Sub

Private Function ReadEntireTextFile(ByVal filePath As String) As String
    Dim fileNumber As Integer
    Dim fileLength As Long

    If Len(Dir$(filePath)) = 0 Then Exit Function

    fileNumber = FreeFile
    Open filePath For Binary Access Read As #fileNumber
    fileLength = LOF(fileNumber)
    If fileLength > 0 Then
        ReadEntireTextFile = Space$(fileLength)
        Get #fileNumber, , ReadEntireTextFile
    End If
    Close #fileNumber
End Function

Private Function ExtractSpringFooter(ByVal existingText As String) As String
    Dim normalizedText As String
    Dim lines As Variant
    Dim lineIndex As Long
    Dim footerStart As Long
    Dim result As String

    normalizedText = NormalizeLineBreaks(existingText)
    lines = Split(normalizedText, vbCrLf)
    footerStart = -1

    For lineIndex = LBound(lines) To UBound(lines)
        If Left$(Trim$(CStr(lines(lineIndex))), 5) = "*****" Then
            footerStart = lineIndex
            Exit For
        End If
    Next lineIndex

    If footerStart >= 0 Then
        For lineIndex = footerStart To UBound(lines)
            If Len(result) > 0 Then result = result & vbCrLf
            result = result & CStr(lines(lineIndex))
        Next lineIndex
        ExtractSpringFooter = result
    Else
        ExtractSpringFooter = _
            "*****" & vbCrLf & _
            "**************" & vbCrLf & _
            "*****add" & vbTab & "set" & vbTab & "of" & vbTab & _
                "lateral" & vbTab & "springs" & vbCrLf & _
            "**" & vbCrLf & _
            "*ELSET," & vbTab & "ELSET=elset_latSprings," & vbTab & _
                "GENERATE" & vbCrLf & _
            ""
    End If
End Function

Private Function UpdateSpringFooterRange(ByVal footerText As String, _
                                         ByVal firstElementNumber As Long, _
                                         ByVal lastElementNumber As Long) As String
    Dim lines As Variant
    Dim lineIndex As Long
    Dim foundElset As Boolean
    Dim result As String

    lines = Split(NormalizeLineBreaks(footerText), vbCrLf)
    For lineIndex = LBound(lines) To UBound(lines) - 1
        If UCase$(Left$(Trim$(CStr(lines(lineIndex))), 6)) = "*ELSET" Then
            lines(lineIndex + 1) = _
                " " & CStr(firstElementNumber) & "," & vbTab & _
                CStr(lastElementNumber) & "," & vbTab & "1"
            foundElset = True
            Exit For
        End If
    Next lineIndex

    If Not foundElset Then
        Err.Raise vbObjectError + 2015, , _
            "The lateral-spring ELSET footer could not be created."
    End If

    result = Join(lines, vbCrLf)
    Do While Len(result) >= 2 And Right$(result, 2) = vbCrLf
        result = Left$(result, Len(result) - 2)
    Loop
    UpdateSpringFooterRange = result
End Function

Private Function NormalizeLineBreaks(ByVal textValue As String) As String
    Dim normalizedText As String

    normalizedText = Replace(textValue, vbCrLf, vbLf)
    normalizedText = Replace(normalizedText, vbCr, vbLf)
    NormalizeLineBreaks = Replace(normalizedText, vbLf, vbCrLf)
End Function

Private Sub WriteEntireTextFileSafely(ByVal filePath As String, _
                                      ByVal textValue As String)
    Dim fileNumber As Integer
    Dim temporaryPath As String
    Dim fileSystem As Object

    temporaryPath = filePath & ".tmp"
    If Len(Dir$(temporaryPath)) > 0 Then Kill temporaryPath

    fileNumber = FreeFile
    Open temporaryPath For Output Access Write As #fileNumber
    Print #fileNumber, textValue;
    Close #fileNumber

    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    fileSystem.CopyFile temporaryPath, filePath, True
    Kill temporaryPath
End Sub

Private Function FindLastSpringOutputRow(ByVal springSheet As Worksheet) As Long
    Dim columnNumber As Long
    Dim candidateRow As Long

    For columnNumber = 2 To 25
        candidateRow = springSheet.Cells(springSheet.Rows.Count, columnNumber).End(xlUp).Row
        If candidateRow > FindLastSpringOutputRow Then
            FindLastSpringOutputRow = candidateRow
        End If
    Next columnNumber

    If FindLastSpringOutputRow < SPRING_OUTPUT_FIRST_ROW Then
        FindLastSpringOutputRow = SPRING_OUTPUT_FIRST_ROW - 1
    End If
End Function

Private Sub ExtendSpringFormatting(ByVal springSheet As Worksheet, _
                                   ByVal springLastRow As Long, _
                                   ByVal previousSpringLastRow As Long)
    If springLastRow <= previousSpringLastRow Then Exit Sub

    springSheet.Range("B7:Y7").Copy
    springSheet.Range("B" & CStr(previousSpringLastRow + 1) & _
                      ":Y" & CStr(springLastRow)).PasteSpecial Paste:=xlPasteFormats
    Application.CutCopyMode = False
End Sub

Private Sub FillSpringMarker(ByVal springSheet As Worksheet, _
                             ByVal columnLetter As String, _
                             ByVal markerValue As Variant, _
                             ByVal springLastRow As Long)
    springSheet.Range(columnLetter & CStr(SPRING_OUTPUT_FIRST_ROW) & _
                      ":" & columnLetter & CStr(springLastRow)).Value2 = markerValue
End Sub

Private Sub ComputeCoordinate(ByVal nodeNumber As Long, _
                              ByRef anchorNodes() As Long, _
                              ByRef anchorX() As Double, _
                              ByRef anchorY() As Double, _
                              ByRef anchorZ() As Double, _
                              ByVal anchorCount As Long, _
                              ByVal curves As Collection, _
                              ByRef xValue As Double, _
                              ByRef yValue As Double, _
                              ByRef zValue As Double)
    Dim exactIndex As Long
    Dim segmentIndex As Long
    Dim curveDefinition As Variant
    Dim startIndex As Long
    Dim endIndex As Long
    Dim ratio As Double
    Dim centerX As Double
    Dim centerY As Double
    Dim radius As Double
    Dim startAngle As Double
    Dim endAngle As Double
    Dim currentAngle As Double

    exactIndex = FindExactAnchorIndex(anchorNodes, anchorCount, nodeNumber)
    If exactIndex >= 0 Then
        xValue = anchorX(exactIndex)
        yValue = anchorY(exactIndex)
        zValue = anchorZ(exactIndex)
        Exit Sub
    End If

    For Each curveDefinition In curves
        If nodeNumber > CLng(curveDefinition(0)) And _
           nodeNumber < CLng(curveDefinition(1)) Then
            startIndex = FindExactAnchorIndex(anchorNodes, anchorCount, CLng(curveDefinition(0)))
            endIndex = FindExactAnchorIndex(anchorNodes, anchorCount, CLng(curveDefinition(1)))
            ratio = (nodeNumber - CLng(curveDefinition(0))) / _
                    (CLng(curveDefinition(1)) - CLng(curveDefinition(0)))
            centerX = CDbl(curveDefinition(2))
            centerY = CDbl(curveDefinition(3))
            radius = Sqr((anchorX(startIndex) - centerX) ^ 2 + _
                         (anchorY(startIndex) - centerY) ^ 2)
            If radius <= 0# Then
                Err.Raise vbObjectError + 2009, , _
                    "A curve has zero radius at node " & CStr(curveDefinition(0)) & "."
            End If

            startAngle = NormalizeAngle(Atan2(anchorY(startIndex) - centerY, _
                                              anchorX(startIndex) - centerX))
            endAngle = NormalizeAngle(Atan2(anchorY(endIndex) - centerY, _
                                            anchorX(endIndex) - centerX))

            If CLng(curveDefinition(4)) = -1 Then
                Do While endAngle >= startAngle
                    endAngle = endAngle - 2# * PI_VALUE
                Loop
            Else
                Do While endAngle <= startAngle
                    endAngle = endAngle + 2# * PI_VALUE
                Loop
            End If

            currentAngle = startAngle + ratio * (endAngle - startAngle)
            xValue = centerX + radius * Cos(currentAngle)
            yValue = centerY + radius * Sin(currentAngle)
            zValue = anchorZ(startIndex) + ratio * (anchorZ(endIndex) - anchorZ(startIndex))
            Exit Sub
        End If
    Next curveDefinition

    segmentIndex = FindSegmentIndex(anchorNodes, anchorCount, nodeNumber)
    If segmentIndex < 0 Then
        Err.Raise vbObjectError + 2010, , _
            "No straight or curve segment contains node " & CStr(nodeNumber) & "."
    End If

    ratio = (nodeNumber - anchorNodes(segmentIndex)) / _
            (anchorNodes(segmentIndex + 1) - anchorNodes(segmentIndex))
    xValue = anchorX(segmentIndex) + ratio * (anchorX(segmentIndex + 1) - anchorX(segmentIndex))
    yValue = anchorY(segmentIndex) + ratio * (anchorY(segmentIndex + 1) - anchorY(segmentIndex))
    zValue = anchorZ(segmentIndex) + ratio * (anchorZ(segmentIndex + 1) - anchorZ(segmentIndex))
End Sub

Private Function FindExactAnchorIndex(ByRef nodes() As Long, _
                                      ByVal itemCount As Long, _
                                      ByVal nodeNumber As Long) As Long
    Dim lowIndex As Long
    Dim highIndex As Long
    Dim middleIndex As Long

    lowIndex = 0
    highIndex = itemCount - 1
    FindExactAnchorIndex = -1

    Do While lowIndex <= highIndex
        middleIndex = (lowIndex + highIndex) \ 2
        If nodes(middleIndex) = nodeNumber Then
            FindExactAnchorIndex = middleIndex
            Exit Function
        ElseIf nodes(middleIndex) < nodeNumber Then
            lowIndex = middleIndex + 1
        Else
            highIndex = middleIndex - 1
        End If
    Loop
End Function

Private Function FindSegmentIndex(ByRef nodes() As Long, _
                                  ByVal itemCount As Long, _
                                  ByVal nodeNumber As Long) As Long
    Dim lowIndex As Long
    Dim highIndex As Long
    Dim middleIndex As Long

    lowIndex = 0
    highIndex = itemCount - 2
    FindSegmentIndex = -1

    Do While lowIndex <= highIndex
        middleIndex = (lowIndex + highIndex) \ 2
        If nodeNumber < nodes(middleIndex) Then
            highIndex = middleIndex - 1
        ElseIf nodeNumber > nodes(middleIndex + 1) Then
            lowIndex = middleIndex + 1
        Else
            FindSegmentIndex = middleIndex
            Exit Function
        End If
    Loop
End Function

Private Function NormalizeAngle(ByVal angleValue As Double) As Double
    Do While angleValue < 0#
        angleValue = angleValue + 2# * PI_VALUE
    Loop
    Do While angleValue >= 2# * PI_VALUE
        angleValue = angleValue - 2# * PI_VALUE
    Loop
    NormalizeAngle = angleValue
End Function

Private Function Atan2(ByVal yValue As Double, ByVal xValue As Double) As Double
    If xValue > 0# Then
        Atan2 = Atn(yValue / xValue)
    ElseIf xValue < 0# And yValue >= 0# Then
        Atan2 = Atn(yValue / xValue) + PI_VALUE
    ElseIf xValue < 0# And yValue < 0# Then
        Atan2 = Atn(yValue / xValue) - PI_VALUE
    ElseIf xValue = 0# And yValue > 0# Then
        Atan2 = PI_VALUE / 2#
    ElseIf xValue = 0# And yValue < 0# Then
        Atan2 = -PI_VALUE / 2#
    Else
        Atan2 = 0#
    End If
End Function

Attribute VB_Name = "OperatingDataWorkflows"
Option Explicit

Private busy As Boolean
Private logRow As Long

Public Sub RunAllOperatingData()
    RunOperatingWorkflow 0
End Sub
Public Sub RunSteadyState()
    RunOperatingWorkflow 1
End Sub
Public Sub RunTemperatureTransients()
    RunOperatingWorkflow 2
End Sub
Public Sub RunPressureTransients()
    RunOperatingWorkflow 3
End Sub
Public Sub GenerateStartupProfiles()
    RunOperatingWorkflow 4
End Sub

Public Sub DefineDataRange()
    Dim chosen As Range, ws As Worksheet, firstRow As Long, firstCol As Long
    On Error Resume Next
    Set chosen = Application.InputBox("Select the raw input range on its sheet. Transients must start at A6; FL6 data at C6; FL6_WD at A2. The last selected row limits processing; transient selection also sets the last profile column.", "Define data range", Type:=8)
    On Error GoTo Failed
    If chosen Is Nothing Then Exit Sub
    Set ws = chosen.Worksheet
    If Not ws.Parent Is ThisWorkbook Then Err.Raise 5, , "Select a range in this workbook."
    firstRow = 6: firstCol = 1
    If ws.Name = "FL6 data" Then
        firstCol = 3
    ElseIf ws.Name = "FL6_WD" Then
        firstRow = 2
    ElseIf Kind(ws.Name) = "" Then
        Err.Raise 5, , "Select a transient input, FL6 data, or FL6_WD sheet."
    End If
    If chosen.Areas.Count <> 1 Or chosen.Row <> firstRow Or chosen.Column <> firstCol Then Err.Raise 5, , "The range must be one rectangle starting at " & ws.Cells(firstRow, firstCol).Address(False, False)
    If chosen.Columns.Count < 2 Then Err.Raise 5, , "Include the distance column and at least one profile column."
    ThisWorkbook.Worksheets("Data Ranges").Cells(ConfigRow(ws.Name), 2).Value = chosen.Address(False, False)
    MsgBox "Range saved for " & ws.Name & ": " & chosen.Address(False, False) & vbCrLf & "Rerun the required workflow. Enter AUTO in Data Ranges column B to restore automatic detection.", vbInformation
    Exit Sub
Failed:
    MsgBox Err.Description, vbExclamation, "Define data range"
End Sub

Private Function ConfigRow(ByVal sheetName As String) As Long
    Dim s As Worksheet, r As Long
    Set s = ThisWorkbook.Worksheets("Data Ranges")
    For r = 2 To s.Cells(s.Rows.Count, 1).End(xlUp).Row
        If CStr(s.Cells(r, 1).Value2) = sheetName Then ConfigRow = r: Exit Function
    Next r
    r = s.Cells(s.Rows.Count, 1).End(xlUp).Row + 1
    s.Cells(r, 1).Value = sheetName: s.Cells(r, 2).Value = "AUTO": s.Cells(r, 3).Value = "User"
    ConfigRow = r
End Function

Private Function SelectedRange(ByVal s As Worksheet) As Range
    Dim a As String, selected As Range, firstRow As Long, firstCol As Long
    a = Trim$(CStr(ThisWorkbook.Worksheets("Data Ranges").Cells(ConfigRow(s.Name), 2).Value2))
    If a <> "" And UCase$(a) <> "AUTO" Then
        Set selected = s.Range(a)
        firstRow = 6: firstCol = 1
        If s.Name = "FL6 data" Then firstCol = 3
        If s.Name = "FL6_WD" Then firstRow = 2
        If selected.Areas.Count <> 1 Or selected.Row <> firstRow Or selected.Column <> firstCol Then Err.Raise 5, , s.Name & ": range must be a rectangle starting at " & s.Cells(firstRow, firstCol).Address(False, False)
        Set SelectedRange = selected
    End If
End Function

Private Function Generated(ByVal s As Worksheet) As Boolean
    Generated = Left$(CStr(ThisWorkbook.Worksheets("Data Ranges").Cells(ConfigRow(s.Name), 3).Value2), 9) = "Generated"
End Function

Private Function HasInput(ByVal s As Worksheet) As Boolean
    Dim r As Long, last As Long
    If Kind(s.Name) = "" Then Exit Function
    If CStr(ThisWorkbook.Worksheets("Data Ranges").Cells(ConfigRow(s.Name), 3).Value2) = "Generated - unavailable" Then Exit Function
    last = LastRow(s, 1)
    For r = 6 To last
        If IsValue(s.Cells(r, 2).Value2) Then HasInput = True: Exit Function
    Next r
End Function

Private Function CaseSheet(ByVal amplitude As String, ByVal dataKind As String) As Worksheet
    Dim s As Worksheet
    For Each s In ThisWorkbook.Worksheets
        If Kind(s.Name) = dataKind Then
            If LCase$(AmpName(s.Name, dataKind)) = LCase$(amplitude) Then Set CaseSheet = s: Exit Function
        End If
    Next s
End Function

' Optional silent mode is for verification; buttons always show completion/errors.
Public Sub RunOperatingWorkflow(ByVal mode As Long, Optional ByVal silent As Boolean = False)
    Dim oldEvents As Boolean, oldScreen As Boolean, oldCalc As XlCalculation
    Dim errNum As Long, errText As String
    If busy Then Exit Sub
    busy = True
    oldEvents = Application.EnableEvents: oldScreen = Application.ScreenUpdating
    oldCalc = Application.Calculation
    On Error GoTo Failed
    Application.EnableEvents = False: Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    StartLog
    If mode < 0 Or mode > 4 Then Err.Raise 5, , "Unknown workflow."
    If mode = 0 Or mode = 1 Then UpdateSteady
    If mode = 0 Or mode = 4 Then BuildMissingStartups
    If mode = 0 Or mode = 2 Or mode = 4 Then UpdateTransients "Temperature"
    If mode = 0 Or mode = 3 Or mode = 4 Then UpdateTransients "Pressure"
    ThisWorkbook.Worksheets("Workflow Controls").Range("C18").Value = Now
    ThisWorkbook.Worksheets("Workflow Controls").Range("C18").NumberFormat = "yyyy-mm-dd hh:mm:ss"
    AddLog "Completed", "Workflow", "Requested outputs refreshed. Review warnings and missing cases below."
    GoTo CleanExit
Failed:
    errNum = Err.Number: errText = Err.Description
    AddLog "ERROR", "Workflow", errText
CleanExit:
    Application.Calculation = oldCalc
    Application.EnableEvents = oldEvents: Application.ScreenUpdating = oldScreen
    busy = False
    ' Silent callers inspect Workflow Log for ERROR; do not open VBA debug dialogs.
    If Not silent Then
        If errNum <> 0 Then
            MsgBox "Workflow stopped: " & errText & vbCrLf & "See Workflow Log. Outputs from earlier stages may already be updated.", vbExclamation
        Else
            MsgBox "Workflow completed. Review Workflow Log for missing cases and input warnings.", vbInformation
        End If
    End If
End Sub

Private Sub StartLog()
    Dim s As Worksheet
    Set s = ThisWorkbook.Worksheets("Workflow Log")
    s.UsedRange.ClearContents
    s.Range("A1:D1").Value = Array("Time", "Status", "Sheet / case", "Details")
    logRow = 2
End Sub
Private Sub AddLog(ByVal status As String, ByVal source As String, ByVal detail As String)
    Dim s As Worksheet
    Set s = ThisWorkbook.Worksheets("Workflow Log")
    If logRow < 2 Then logRow = 2
    s.Cells(logRow, 1).Value = Now
    s.Cells(logRow, 1).NumberFormat = "hh:mm:ss"
    s.Cells(logRow, 2).Value = status
    s.Cells(logRow, 3).Value = source
    s.Cells(logRow, 4).Value = detail
    logRow = logRow + 1
End Sub
Private Function LastRow(ByVal s As Worksheet, ByVal col As Long) As Long
    Dim chosen As Range
    LastRow = s.Cells(s.Rows.Count, col).End(xlUp).Row
    If Kind(s.Name) <> "" Or s.Name = "FL6 data" Or s.Name = "FL6_WD" Then
        Set chosen = SelectedRange(s)
        If Not chosen Is Nothing Then LastRow = Application.Min(LastRow, chosen.Row + chosen.Rows.Count - 1)
    End If
End Function
Private Function LastCol(ByVal s As Worksheet) As Long
    Dim f As Range
    Set f = s.Cells.Find("*", s.Cells(1, 1), xlFormulas, xlPart, xlByColumns, xlPrevious)
    If Not f Is Nothing Then LastCol = f.Column
End Function
Private Function ColName(ByVal col As Long) As String
    ColName = Split(ThisWorkbook.Worksheets(1).Cells(1, col).Address, "$")(1)
End Function
Private Function Ref(ByVal s As Worksheet, ByVal row As Long, ByVal col As Long) As String
    Ref = "'" & Replace(s.Name, "'", "''") & "'!" & s.Cells(row, col).Address(False, False)
End Function
Private Function IsValue(ByVal v As Variant) As Boolean
    If IsError(v) Or IsEmpty(v) Then Exit Function
    If VarType(v) = vbString And Len(Trim$(v)) = 0 Then Exit Function
    IsValue = IsNumeric(v)
End Function

Private Function Kind(ByVal name As String) As String
    Dim n As String
    n = LCase$(name)
    If Right$(n, 8) = "pressure" Then Kind = "Pressure"
    If Right$(n, 4) = "temp" Or Right$(n, 11) = "temperature" Then Kind = "Temperature"
    If InStr(n, "_cd_") = 0 And InStr(n, "_hu_") = 0 Then Kind = ""
End Function
Private Function AmpName(ByVal name As String, ByVal dataKind As String) As String
    Dim n As Long
    If dataKind = "Pressure" Then
        n = 8
    ElseIf LCase$(Right$(name, 11)) = "temperature" Then
        n = 11
    Else
        n = 4
    End If
    AmpName = Left$(name, Len(name) - n)
    If Right$(AmpName, 1) = "_" Then AmpName = Left$(AmpName, Len(AmpName) - 1)
End Function

' Raw data ends at the first fully empty column after A. Conversion blocks
' are identified by their Distance header and kept separate from raw inputs.
Private Sub Layout(ByVal s As Worksheet, ByRef rawEnd As Long, ByRef endRow As Long, ByRef conv As Long)
    Dim c As Long, limit As Long, chosen As Range
    endRow = LastRow(s, 1)
    If endRow < 6 Then Err.Raise 5, , s.Name & ": no raw data at row 6."
    limit = LastCol(s) + 1
    For c = 2 To limit
        If Application.CountA(s.Range(s.Cells(2, c), s.Cells(endRow, c))) = 0 Then Exit For
    Next c
    rawEnd = c - 1
    Set chosen = SelectedRange(s)
    If Not chosen Is Nothing Then
        If chosen.Row <> 6 Or chosen.Column <> 1 Or chosen.Areas.Count <> 1 Then Err.Raise 5, , s.Name & ": range must start at A6."
        rawEnd = chosen.Columns.Count
        If rawEnd >= c Then Err.Raise 5, , s.Name & ": selected profiles include the blank separator or conversion block."
    End If
    If rawEnd < 2 Then Err.Raise 5, , s.Name & ": no profile columns before the blank separator."
    conv = 0
    For c = rawEnd + 2 To limit
        If LCase$(Trim$(CStr(s.Cells(4, c).Value2))) = "distance" Then
            conv = c: Exit For
        End If
    Next c
    If conv = 0 Then conv = limit + 1
    If conv + rawEnd - 1 > s.Columns.Count Then Err.Raise 5, , "Too many profile columns on " & s.Name
End Sub

Private Sub ValidateRaw(ByVal s As Worksheet, ByVal rawEnd As Long, ByVal endRow As Long, ByVal timeRow As Long)
    Dim a As Variant, r As Long, c As Long, badTime As Boolean
    a = s.Range(s.Cells(6, 1), s.Cells(endRow, rawEnd)).Value2
    For r = 1 To UBound(a, 1)
        For c = 1 To rawEnd
            If Not IsValue(a(r, c)) Then Err.Raise 5, , s.Name & ": missing or nonnumeric input at " & s.Cells(r + 5, c).Address(False, False)
        Next c
        If r > 1 Then
            If CDbl(a(r, 1)) <= CDbl(a(r - 1, 1)) Then Err.Raise 5, , s.Name & ": distance must increase strictly down column A."
        End If
    Next r
    If a(1, 1) < 0 Then Err.Raise 5, , s.Name & ": negative starting KP is incompatible with the KP=0 template boundary."
    For c = 2 To rawEnd
        If Len(Trim$(CStr(s.Cells(3, c).Value2))) = 0 Then Err.Raise 5, , s.Name & ": missing transient name in row 3, column " & ColName(c)
        If Len(Trim$(CStr(s.Cells(timeRow, c).Value2))) = 0 Then Err.Raise 5, , s.Name & ": missing time in row " & timeRow & ", column " & ColName(c)
        If c > 2 And IsValue(s.Cells(timeRow, c).Value2) And IsValue(s.Cells(timeRow, c - 1).Value2) Then
            If CDbl(s.Cells(timeRow, c).Value2) <= CDbl(s.Cells(timeRow, c - 1).Value2) Then badTime = True
        End If
    Next c
    If badTime Then AddLog "WARNING", s.Name, "Time values are not strictly increasing. Copied exactly; correct the raw time row before using the profiles."
End Sub

Private Sub ConvertRaw(ByVal s As Worksheet, ByVal dataKind As String, ByVal rawEnd As Long, ByVal endRow As Long, ByVal conv As Long, ByVal timeRow As Long)
    Dim c As Long, oldEnd As Long, oldCol As Long, formulas() As Variant
    oldEnd = s.UsedRange.Row + s.UsedRange.Rows.Count - 1
    oldCol = LastCol(s)
    If oldCol >= conv Then s.Range(s.Cells(2, conv), s.Cells(Application.Max(6, oldEnd), oldCol)).ClearContents
    s.Cells(4, conv).Value = "Distance": s.Cells(5, conv).Value = "[m]"
    For c = 2 To rawEnd
        s.Cells(2, conv + c - 1).Formula = "=" & s.Cells(timeRow, c).Address(False, False)
        s.Cells(3, conv + c - 1).Formula = "=" & s.Cells(3, c).Address(False, False)
        s.Cells(4, conv + c - 1).Value = dataKind
        If dataKind = "Temperature" Then s.Cells(5, conv + c - 1).Value = "[C]" Else s.Cells(5, conv + c - 1).Value = "[bar]"
    Next c
    ReDim formulas(1 To 1, 1 To rawEnd)
    formulas(1, 1) = "=A6*0.3048"
    For c = 2 To rawEnd
        If dataKind = "Temperature" Then
            formulas(1, c) = "=(" & ColName(c) & "6-32)*5/9"
        Else
            formulas(1, c) = "=" & ColName(c) & "6*0.068948"
        End If
    Next c
    s.Range(s.Cells(6, conv), s.Cells(6, conv + rawEnd - 1)).Formula = formulas
    s.Range(s.Cells(6, conv), s.Cells(endRow, conv + rawEnd - 1)).FillDown
    s.Range(s.Cells(6, conv), s.Cells(endRow, conv + rawEnd - 1)).NumberFormat = "0.0000"
    s.Range(s.Cells(4, conv), s.Cells(4, conv + rawEnd - 1)).EntireColumn.ColumnWidth = 14
    s.Range(s.Cells(2, conv), s.Cells(endRow, conv + rawEnd - 1)).Calculate
End Sub

Private Sub UpdateTransients(ByVal dataKind As String)
    Dim s As Worksheet, dest As Worksheet, first As Worksheet, cases As New Collection
    Dim widths As Object, ends As Object, starts As Object, seen As Object
    Dim rawEnd As Long, endRow As Long, conv As Long, timeRow As Long
    Dim c As Long, r As Long, outCol As Long, n As Long, addZero As Long
    Dim refKP As Variant, kp As Variant, amp As String, life As Variant, op As Variant, phase As Variant
    Dim totalCols As Long, oldRows As Long, oldCols As Long
    Set widths = CreateObject("Scripting.Dictionary"): Set ends = CreateObject("Scripting.Dictionary")
    Set starts = CreateObject("Scripting.Dictionary"): Set seen = CreateObject("Scripting.Dictionary")
    timeRow = CLng(ThisWorkbook.Worksheets("Workflow Controls").Range("C17").Value2)
    If timeRow <> 2 And timeRow <> 5 Then Err.Raise 5, , "Set the transient time source row to 2 or 5 on Workflow Controls."
    totalCols = 1
    For Each s In ThisWorkbook.Worksheets
        If Kind(s.Name) = dataKind And HasInput(s) Then
            Layout s, rawEnd, endRow, conv
            ValidateRaw s, rawEnd, endRow, timeRow
            amp = AmpName(s.Name, dataKind)
            If seen.Exists(LCase$(amp)) Then Err.Raise 5, , "Duplicate amplitude name: " & amp
            seen.Add LCase$(amp), True
            cases.Add s: widths.Add s.Name, rawEnd: ends.Add s.Name, endRow: starts.Add s.Name, conv
            kp = s.Range("A6:A" & endRow).Value2
            If first Is Nothing Then
                Set first = s: refKP = kp
            Else
                If UBound(kp, 1) <> UBound(refKP, 1) Then Err.Raise 5, , s.Name & ": KP row count differs from " & first.Name & ". The combined output requires the same KP grid."
                For r = 1 To UBound(kp, 1)
                    If Abs(CDbl(kp(r, 1)) - CDbl(refKP(r, 1))) > 0.000001 Then Err.Raise 5, , s.Name & ": KP grid differs from " & first.Name & " at row " & r + 5
                Next r
            End If
            totalCols = totalCols + rawEnd - 1
        End If
    Next s
    If totalCols > 16384 Then Err.Raise 5, , "The combined transient output exceeds Excel's column limit."
    Set dest = ThisWorkbook.Worksheets(dataKind & "Transient")
    oldRows = dest.UsedRange.Row + dest.UsedRange.Rows.Count - 1
    oldCols = LastCol(dest)
    ' Validate all cases before replacing the output. Keep template styles.
    dest.Range(dest.Cells(1, 2), dest.Cells(1, Application.Max(2, oldCols))).UnMerge
    dest.Range(dest.Cells(1, 2), dest.Cells(Application.Max(6, oldRows), Application.Max(2, oldCols))).ClearContents
    dest.Range("A6:A" & Application.Max(6, oldRows)).ClearContents
    If cases.Count = 0 Then
        AddLog "MISSING", dataKind, "No matching transient input tabs. Old generated output cleared."
        Exit Sub
    End If
    n = UBound(refKP, 1): addZero = 0
    If CDbl(refKP(1, 1)) > 0 Then addZero = 1
    outCol = 2
    For Each s In cases
        rawEnd = widths(s.Name): endRow = ends(s.Name): conv = starts(s.Name)
        ConvertRaw s, dataKind, rawEnd, endRow, conv, timeRow
        dest.Range("B2:B6").Copy
        dest.Range(dest.Cells(2, outCol), dest.Cells(6, outCol + rawEnd - 2)).PasteSpecial xlPasteFormats
        dest.Range("B7").Copy
        dest.Range(dest.Cells(7, outCol), dest.Cells(5 + n + addZero, outCol + rawEnd - 2)).PasteSpecial xlPasteFormats
        With dest.Range(dest.Cells(1, outCol), dest.Cells(1, outCol + rawEnd - 2))
            .Font.Name = dest.Range("B1").Font.Name
            .Font.Size = dest.Range("B1").Font.Size
            .Font.Color = dest.Range("B1").Font.Color
            .Interior.Color = dest.Range("B1").Interior.Color
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
        dest.Range(dest.Cells(1, outCol), dest.Cells(1, outCol + rawEnd - 2)).Merge
        dest.Cells(1, outCol).Value = AmpName(s.Name, dataKind)
        dest.Range(dest.Cells(2, outCol), dest.Cells(2, outCol + rawEnd - 2)).Value2 = s.Range(s.Cells(3, 2), s.Cells(3, rawEnd)).Value2
        dest.Range(dest.Cells(3, outCol), dest.Cells(3, outCol + rawEnd - 2)).Value2 = s.Range(s.Cells(timeRow, 2), s.Cells(timeRow, rawEnd)).Value2
        dest.Range(dest.Cells(4, outCol), dest.Cells(4, outCol + rawEnd - 2)).Value = dataKind
        If dataKind = "Temperature" Then
            dest.Range(dest.Cells(5, outCol), dest.Cells(5, outCol + rawEnd - 2)).Value = "[C]"
        Else
            dest.Range(dest.Cells(5, outCol), dest.Cells(5, outCol + rawEnd - 2)).Value = "[bar]"
        End If
        dest.Range(dest.Cells(6 + addZero, outCol), dest.Cells(5 + n + addZero, outCol + rawEnd - 2)).Value2 = s.Range(s.Cells(6, conv + 1), s.Cells(endRow, conv + rawEnd - 1)).Value2
        If addZero = 1 Then dest.Range(dest.Cells(6, outCol), dest.Cells(6, outCol + rawEnd - 2)).Value2 = s.Range(s.Cells(6, conv + 1), s.Cells(6, conv + rawEnd - 1)).Value2
        AddLog "Updated", s.Name, CStr(n) & " KP rows; " & rawEnd - 1 & " profiles; output " & ColName(outCol) & ":" & ColName(outCol + rawEnd - 2)
        outCol = outCol + rawEnd - 1
    Next s
    conv = starts(first.Name)
    dest.Range(dest.Cells(6 + addZero, 1), dest.Cells(5 + n + addZero, 1)).Value2 = first.Range(first.Cells(6, conv), first.Cells(5 + n, conv)).Value2
    If addZero = 1 Then dest.Range("A6").Value2 = 0
    Application.CutCopyMode = False
    For Each life In Array("Early", "Middle", "Late")
        If LifeSupplied(CStr(life), dataKind) Then
        For Each op In Array("FCD", "HCD", "PCD")
            For Each phase In Array("CD", "HU")
                amp = LCase$(op & "_" & phase & "_" & life)
                If Not seen.Exists(amp) Then AddLog "MISSING", op & "_" & phase & "_" & life & " " & dataKind, "No input tab supplied. No data invented. Add the named case tab to include it on the next run."
            Next phase
        Next op
        End If
    Next life
End Sub

Private Function LifeSupplied(ByVal life As String, ByVal dataKind As String) As Boolean
    Dim s As Worksheet
    For Each s In ThisWorkbook.Worksheets
        If Kind(s.Name) = dataKind And InStr(1, s.Name, "_" & life, vbTextCompare) > 0 Then
            If Not Generated(s) And HasInput(s) Then LifeSupplied = True: Exit Function
        End If
    Next s
End Function

Private Function SteadyLifeSupplied(ByVal life As String) As Boolean
    Dim s As Worksheet, cols As Variant, col As Variant, r As Long, last As Long
    Set s = ThisWorkbook.Worksheets("FL6 data")
    If life = "Middle" Then cols = Array(10, 11, 12, 27, 28, 29, 45, 47, 49) Else cols = Array(16, 17, 18, 32, 33, 34, 52, 54, 56)
    For Each col In cols
        last = LastRow(s, CLng(col))
        If last >= 6 Then
            If Application.Count(s.Range(s.Cells(6, CLng(col)), s.Cells(last, CLng(col)))) > 0 Then SteadyLifeSupplied = True: Exit Function
        End If
    Next col
    If life = "Late" Then
        For Each col In Array("C35", "C36")
            If Len(Trim$(CStr(ThisWorkbook.Worksheets("Workflow Controls").Range(CStr(col)).Value2))) > 0 Then
                r = s.Range(CStr(ThisWorkbook.Worksheets("Workflow Controls").Range(CStr(col)).Value2) & "1").Column
                last = LastRow(s, r)
                If last >= 6 Then
                    If Application.Count(s.Range(s.Cells(6, r), s.Cells(last, r))) > 0 Then SteadyLifeSupplied = True: Exit Function
                End If
            End If
        Next col
    End If
End Function

Private Sub BuildMissingStartups()
    Dim life As Variant, typ As Variant, op As Variant, base As Worksheet, target As Worksheet
    For Each typ In Array("Temperature", "Pressure")
        For Each life In Array("Early", "Middle", "Late")
            Set base = CaseSheet("FCD_HU_" & life, CStr(typ))
            For Each op In Array("HCD", "PCD")
                Set target = CaseSheet(op & "_HU_" & life, CStr(typ))
                If Not target Is Nothing Then
                    If Not Generated(target) And HasInput(target) Then GoTo NextCase
                End If
                If base Is Nothing Then
                    If Not target Is Nothing Then
                        If Generated(target) Then ThisWorkbook.Worksheets("Data Ranges").Cells(ConfigRow(target.Name), 3).Value = "Generated - unavailable"
                    End If
                    GoTo NextCase
                End If
                If Not HasInput(base) Then
                    If Not target Is Nothing Then
                        If Generated(target) Then ThisWorkbook.Worksheets("Data Ranges").Cells(ConfigRow(target.Name), 3).Value = "Generated - unavailable"
                    End If
                    GoTo NextCase
                End If
                MakeStartup base, target, CStr(op), CStr(life), CStr(typ)
NextCase:
            Next op
        Next life
    Next typ
End Sub

Private Sub MakeStartup(ByVal base As Worksheet, ByVal target As Worksheet, ByVal op As String, ByVal life As String, ByVal typ As String)
    Dim src As Worksheet, kpCol As Long, valCol As Long, idx As Long, lastSS As Long
    Dim rawEnd As Long, endRow As Long, conv As Long, timeRow As Long, extra As Long, count As Long
    Dim r As Long, c As Long, j As Long, k As Long, value As Double, x As Double, y As Double
    Dim ssX As Variant, ssY As Variant, a As Variant, result() As Variant, newName As String
    Dim pCols As Variant, tCols As Variant, xCols As Variant, rc As Long
    On Error GoTo Unavailable
    newName = op & "_HU_" & life
    If typ = "Temperature" Then newName = newName & "Temp" Else newName = newName & "Pressure"
    Set src = ThisWorkbook.Worksheets("FL6 data")
    idx = 0: If life = "Middle" Then idx = 1
    If life = "Late" Then idx = 2
    If op = "PCD" Then
        xCols = Array(21, 26, 31): tCols = Array(22, 27, 32): pCols = Array(23, 28, 33)
    Else
        xCols = Array(37, 44, 51): tCols = Array(38, 45, 52): pCols = Array(40, 47, 54)
    End If
    kpCol = xCols(idx)
    If typ = "Temperature" Then
        valCol = tCols(idx)
    Else
        valCol = pCols(idx)
        If op = "HCD" Then kpCol = kpCol + 2
    End If
    lastSS = Application.Min(LastRow(src, kpCol), LastRow(src, valCol))
    If lastSS < 7 Then Err.Raise 5, , "At least two steady-state KP/value rows are required."
    ssX = src.Range(src.Cells(6, kpCol), src.Cells(lastSS, kpCol)).Value2
    ssY = src.Range(src.Cells(6, valCol), src.Cells(lastSS, valCol)).Value2
    For r = 1 To UBound(ssX, 1)
        If Not IsValue(ssX(r, 1)) Or Not IsValue(ssY(r, 1)) Then Err.Raise 5, , "Missing steady-state KP/value at row " & r + 5
        If r > 1 Then
            If ssX(r, 1) <= ssX(r - 1, 1) Then Err.Raise 5, , "Steady-state KPs must increase strictly."
        End If
    Next r
    Layout base, rawEnd, endRow, conv
    timeRow = CLng(ThisWorkbook.Worksheets("Workflow Controls").Range("C17").Value2)
    ValidateRaw base, rawEnd, endRow, timeRow
    a = base.Range(base.Cells(6, 1), base.Cells(endRow, rawEnd)).Value2
    If a(1, 1) < ssX(1, 1) - 0.000001 Or a(UBound(a, 1), 1) > ssX(UBound(ssX, 1), 1) + 0.000001 Then Err.Raise 5, , "Steady-state coverage " & ssX(1, 1) & " to " & ssX(UBound(ssX, 1), 1) & " ft does not cover FCD input. Define compatible ranges or supply the missing steady-state data."
    If Not IsValue(base.Cells(timeRow, 2).Value2) Then Err.Raise 5, , "The first FCD startup time must be numeric."
    extra = 1: If CDbl(base.Cells(timeRow, 2).Value2) = 0 Then extra = 0
    count = rawEnd + extra
    ReDim result(1 To UBound(a, 1), 1 To count)
    j = 1
    For r = 1 To UBound(a, 1)
        x = CDbl(a(r, 1))
        Do While j < UBound(ssX, 1) - 1
            If ssX(j + 1, 1) >= x Then Exit Do
            j = j + 1
        Loop
        y = CDbl(ssY(j, 1)) + (x - CDbl(ssX(j, 1))) * (CDbl(ssY(j + 1, 1)) - CDbl(ssY(j, 1))) / (CDbl(ssX(j + 1, 1)) - CDbl(ssX(j, 1)))
        result(r, 1) = x: result(r, 2) = y
        For c = 2 To rawEnd
            k = c + extra
            If k > 2 Then result(r, k) = Application.Max(y, CDbl(a(r, c)))
        Next c
    Next r
    If target Is Nothing Then
        Set target = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        target.Name = newName
    End If
    target.UsedRange.ClearContents
    target.Range("A1").Value = "Generated from " & base.Name & " and FL6 data " & ColName(kpCol) & "/" & ColName(valCol) & ". Rerun Generate PCD/HCD Startup Profiles to refresh."
    target.Range("A4").Value = "Distance [ft]": target.Range("A5").Value = "offset"
    target.Range("B2").Value2 = 0: target.Range("B3").Value = "0 m": target.Range("B5").Value2 = 0
    For c = 2 To rawEnd
        k = c + extra
        If k > 2 Then
            target.Cells(2, k).Value2 = base.Cells(2, c).Value2
            target.Cells(3, k).Value2 = base.Cells(3, c).Value2
            target.Cells(5, k).Value2 = base.Cells(5, c).Value2
        End If
    Next c
    If typ = "Temperature" Then target.Range(target.Cells(4, 2), target.Cells(4, count)).Value = "Temperature [F]" Else target.Range(target.Cells(4, 2), target.Cells(4, count)).Value = "Pressure [psi]"
    target.Range("A6").Resize(UBound(result, 1), count).Value2 = result
    target.Range("A1").Resize(5, count).EntireColumn.ColumnWidth = 16
    target.Range("A6").Resize(UBound(result, 1), count).NumberFormat = "0.0000"
    rc = ConfigRow(target.Name)
    ThisWorkbook.Worksheets("Data Ranges").Cells(rc, 2).Value = "AUTO"
    ThisWorkbook.Worksheets("Data Ranges").Cells(rc, 3).Value = "Generated"
    ThisWorkbook.Worksheets("Data Ranges").Cells(rc, 4).Value = base.Name & "; steady " & ColName(kpCol) & "/" & ColName(valCol)
    AddLog "Generated", newName, "Start line = " & op & " steady state. Later profiles = MAX(start line, FCD). " & count - 1 & " profiles; " & UBound(a, 1) & " KP rows."
    Exit Sub
Unavailable:
    AddLog "MISSING", newName, Err.Description
    If Not target Is Nothing Then
        If Generated(target) Then ThisWorkbook.Worksheets("Data Ranges").Cells(ConfigRow(target.Name), 3).Value = "Generated - unavailable"
    End If
End Sub

Private Sub FillColumn(ByVal s As Worksheet, ByVal col As Long, ByVal first As Long, ByVal last As Long, ByVal formula As String)
    s.Cells(first, col).Formula = formula
    If last > first Then s.Range(s.Cells(first, col), s.Cells(last, col)).FillDown
End Sub
Private Sub TrimBlock(ByVal s As Worksheet, ByVal col1 As Long, ByVal col2 As Long, ByVal last As Long)
    Dim old As Long
    old = s.UsedRange.Row + s.UsedRange.Rows.Count - 1
    If old > last Then s.Range(s.Cells(last + 1, col1), s.Cells(old, col2)).ClearContents
End Sub

Private Sub UpdateSteady()
    Dim src As Worksheet, t As Worksheet, p As Worksheet, d As Worksheet, wd As Worksheet
    Dim endRaw As Long, last As Long, wdLast As Long, c As Long, i As Long, rr As Long
    Dim cols As Variant, funcs As Variant, f As String, col As Variant, tCols As Variant, pCols As Variant
    Set src = ThisWorkbook.Worksheets("FL6 data")
    Set t = ThisWorkbook.Worksheets("Temperature"): Set p = ThisWorkbook.Worksheets("Pressure")
    Set d = ThisWorkbook.Worksheets("Density"): Set wd = ThisWorkbook.Worksheets("FL6_WD")
    endRaw = LastRow(src, 3)
    If endRaw < 6 Then Err.Raise 5, , "FL6 data: no steady-state data in C6 onward."
    For rr = 6 To endRaw
        If Not IsValue(src.Cells(rr, 3).Value2) Then Err.Raise 5, , "Missing steady-state distance in FL6 data!C" & rr
        If rr > 6 Then
            If CDbl(src.Cells(rr, 3).Value2) <= CDbl(src.Cells(rr - 1, 3).Value2) Then Err.Raise 5, , "Steady-state distances must increase."
        End If
    Next rr
    last = endRaw - 1
    For Each col In Array(9, 15)
        rr = LastRow(src, CLng(col) - 1)
        FillColumn src, CLng(col), 6, rr, "=" & ColName(CLng(col) - 1) & "6*5280"
        TrimBlock src, CLng(col), CLng(col), rr
    Next col
    src.Calculate
    ' Extend the template output table and retain its installation/design values.
    t.Range("A5:Q5").Copy
    t.Range("A5:Q" & last).PasteSpecial xlPasteAll
    p.Range("A5:U5").Copy
    p.Range("A5:U" & last).PasteSpecial xlPasteAll
    FillColumn t, 1, 5, last, "='FL6 data'!C6*0.3048"
    FillColumn p, 1, 5, last, "=Temperature!A5"
    t.Range("A4").Value2 = 0: p.Range("A4").Value2 = 0
    tCols = Array("D", "J", "P", "V", "AA", "AF", "AL", "AS", "AZ", "BN")
    pCols = Array("E", "K", "Q", "W", "AB", "AG", "AN", "AU", "BB", "BK")
    For i = 0 To 9
        rr = 6: If i = 9 Then rr = 8
        f = "'FL6 data'!" & tCols(i) & rr
        FillColumn t, 7 + i, 5, last, "=IF(ISNUMBER(" & f & "),(" & f & "-32)*5/9,NA())"
        f = "'FL6 data'!" & pCols(i) & rr
        FillColumn p, 11 + i, 5, last, "=IF(ISNUMBER(" & f & ")," & f & "*0.068948,NA())"
    Next i
    FillColumn t, 17, 5, last, "=P5"
    FillColumn p, 21, 5, last, "=T5"
    For c = 7 To 17: t.Cells(4, c).Formula = "=" & ColName(c) & "5": Next c
    For c = 11 To 21: p.Cells(4, c).Formula = "=" & ColName(c) & "5": Next c
    ' Missing late FCD is explicit. Optional source-column controls allow real data.
    t.Range("Q1:Q3").Copy t.Range("R1:R3"): t.Range("R1").Value = "FCD_Lat_Temp"
    p.Range("U1:U3").Copy p.Range("V1:V3"): p.Range("V1").Value = "FCD_Lat"
    FillLate src, t, 18, last, "C35", True
    FillLate src, p, 22, last, "C36", False
    TrimBlock t, 1, 18, last: TrimBlock p, 1, 22, last
    ' Landing-temperature interpolation retains the template's Y11:Z15 anchors.
    FillColumn t, 29, 4, last, "=A4"
    FillColumn t, 30, 4, last, "=MIN(MATCH(AC4,$Y$11:$Y$15,1),ROWS($Y$11:$Y$15)-1)"
    FillColumn t, 31, 4, last, "=AD4+1"
    FillColumn t, 32, 4, last, "=INDEX($Y$11:$Y$15,AD4)"
    FillColumn t, 33, 4, last, "=INDEX($Y$11:$Y$15,AE4)"
    FillColumn t, 34, 4, last, "=INDEX($Z$11:$Z$15,AD4)"
    FillColumn t, 35, 4, last, "=INDEX($Z$11:$Z$15,AE4)"
    FillColumn t, 36, 4, last, "=IF(OR(AC4<$Y$11,AC4>$Y$15),NA(),FORECAST(AC4,AH4:AI4,AF4:AG4))"
    FillColumn t, 2, 4, last, "=AJ4"
    FillColumn t, 19, 4, last, "=A4*3.281"
    FillColumn t, 20, 4, last, "=$S$" & last & "-S4"
    FillColumn t, 21, 4, last, "=G4"
    FillColumn t, 22, 4, last, "=Pressure!K4"
    TrimBlock t, 19, 22, last: TrimBlock t, 29, 36, last
    wdLast = LastRow(wd, 1)
    If wdLast < 3 Then Err.Raise 5, , "FL6_WD requires at least two raw water-depth rows."
    For rr = 2 To wdLast
        If Not IsValue(wd.Cells(rr, 1).Value2) Or Not IsValue(wd.Cells(rr, 2).Value2) Then Err.Raise 5, , "Missing water-depth input at row " & rr
        If rr > 2 Then
            If wd.Cells(rr, 1).Value2 <= wd.Cells(rr - 1, 1).Value2 Then Err.Raise 5, , "Water-depth KP must increase strictly."
        End If
    Next rr
    wd.Range("C2:K2").Copy
    wd.Range("C2:K" & wdLast).PasteSpecial xlPasteAll
    TrimBlock wd, 3, 11, wdLast
    FillColumn wd, 13, 3, last - 1, "=Temperature!A4"
    FillColumn wd, 14, 3, last - 1, "=IF(OR(M3<$C$2,M3>$C$" & wdLast & "),NA(),MIN(MATCH(M3,$C$2:$C$" & wdLast & ",1),ROWS($C$2:$C$" & wdLast & ")-1))"
    FillColumn wd, 15, 3, last - 1, "=N3+1"
    FillColumn wd, 16, 3, last - 1, "=INDEX($C$2:$C$" & wdLast & ",N3)"
    FillColumn wd, 17, 3, last - 1, "=INDEX($C$2:$C$" & wdLast & ",O3)"
    FillColumn wd, 18, 3, last - 1, "=INDEX($D$2:$D$" & wdLast & ",N3)"
    FillColumn wd, 19, 3, last - 1, "=INDEX($D$2:$D$" & wdLast & ",O3)"
    wd.Range("T3:AA3").Copy
    wd.Range("T3:AA" & last - 1).PasteSpecial xlPasteAll
    TrimBlock wd, 13, 27, last - 1
    cols = Array("F", "L", "R", "X", "AC", "AH", "AP", "AW", "BD")
    funcs = Array("MIN", "MAX", "AVERAGE")
    For i = 0 To 8
        rr = LastRow(src, src.Range(cols(i) & "1").Column)
        For c = 0 To 2
            f = "'FL6 data'!" & cols(i) & "$6:" & cols(i) & "$" & Application.Max(6, rr)
            d.Cells(4 + c, 5 + i).Formula = "=IF(COUNT(" & f & ")=0,NA()," & funcs(c) & "(" & f & ")*16.018)"
        Next c
    Next i
    t.Range("A4:V" & last).NumberFormat = "0.00"
    p.Range("A4:V" & last).NumberFormat = "0.00"
    t.Range("A1:R3").WrapText = True: p.Range("A1:V3").WrapText = True
    t.Range("A1:R1").EntireColumn.ColumnWidth = 20
    p.Range("A1:V1").EntireColumn.ColumnWidth = 20
    t.Rows(1).RowHeight = 32: p.Rows(1).RowHeight = 32
    d.Range("A1:M3").WrapText = True
    d.Range("A1:M1").EntireColumn.ColumnWidth = 20
    d.Rows(1).RowHeight = 32: d.Rows(2).RowHeight = 28
    d.Range("B4:M6").NumberFormat = "0.00"
    wd.Range("A1:K1").WrapText = True: wd.Range("M1:AA1").WrapText = True
    wd.Range("A1:K1").EntireColumn.ColumnWidth = 18
    wd.Range("M1:AA1").EntireColumn.ColumnWidth = 18
    wd.Rows(1).RowHeight = 32
    wd.Range("A2:K" & wdLast).NumberFormat = "0.00"
    wd.Range("M3:AA" & last - 1).NumberFormat = "0.00"
    ClearAbsentSteadyLife "Middle", t, p, d, last
    ClearAbsentSteadyLife "Late", t, p, d, last
    t.Calculate: wd.Calculate: p.Calculate: t.Calculate: d.Calculate
    AddLog "Updated", "Steady state", endRaw - 5 & " input rows; " & last - 3 & " output rows including KP=0."
    AddLog "WARNING", "FCD early/middle", "Template middle-life FCD copies early-life FCD (Temperature Q=P; Pressure U=T). This original assumption is retained."
    AddLog "WARNING", "Density FCD headings", "Template FCD density columns use HCD raw data AP/AW/BD. Original mappings are retained."
    If LastRow(src, 66) < endRaw + 2 Or LastRow(src, 63) < endRaw + 2 Then AddLog "MISSING", "FCD steady profiles", "BN/BK source rows end before the full steady-state KP grid under the original row mapping. Missing tail values show #N/A."
    If CDbl(src.Cells(endRaw, 3).Value2) > CDbl(wd.Cells(wdLast, 1).Value2) Then AddLog "MISSING", "FL6_WD", "Raw depth ends at " & wd.Cells(wdLast, 1).Value2 & " ft; steady route ends at " & src.Cells(endRaw, 3).Value2 & " ft. Out-of-coverage pressure values show #N/A. Extend A:B and rerun."
    AddLog "NOTE", "FL6 data", "Column C is treated as feet, following the supplied formulas and values (its original header said miles)."
    Application.CutCopyMode = False
End Sub

Private Sub FillLate(ByVal src As Worksheet, ByVal dest As Worksheet, ByVal targetCol As Long, ByVal last As Long, ByVal setting As String, ByVal isTemp As Boolean)
    Dim sourceCol As String, rr As Long, f As String
    sourceCol = Trim$(CStr(ThisWorkbook.Worksheets("Workflow Controls").Range(setting).Value2))
    If Not SteadyLifeSupplied("Late") And sourceCol = "" Then
        dest.Range(dest.Cells(4, targetCol), dest.Cells(last, targetCol)).ClearContents
        Exit Sub
    End If
    If sourceCol = "" Then
        FillColumn dest, targetCol, 4, last, "=NA()"
        AddLog "MISSING", "FCD late " & dest.Name, "No source column supplied. Enter a FL6 data column letter in Workflow Controls!" & setting & " when available."
    Else
        rr = CLng(ThisWorkbook.Worksheets("Workflow Controls").Range("C37").Value2)
        If rr < 6 Then Err.Raise 5, , "Late FCD first data row must be at least 6."
        f = "'FL6 data'!" & ColName(src.Range(sourceCol & "1").Column) & rr
        If isTemp Then f = "=IF(ISNUMBER(" & f & "),(" & f & "-32)*5/9,NA())" Else f = "=IF(ISNUMBER(" & f & ")," & f & "*0.068948,NA())"
        FillColumn dest, targetCol, 5, last, f
        dest.Cells(4, targetCol).Formula = "=" & ColName(targetCol) & "5"
    End If
End Sub

Private Sub ClearAbsentSteadyLife(ByVal life As String, ByVal t As Worksheet, ByVal p As Worksheet, ByVal d As Worksheet, ByVal last As Long)
    Dim tc As Variant, pc As Variant, dc As Variant, c As Variant
    If SteadyLifeSupplied(life) Then Exit Sub
    If life = "Middle" Then
        tc = Array(8, 11, 14, 17): pc = Array(12, 15, 18, 21): dc = Array(6, 9, 12)
    Else
        tc = Array(9, 12, 15, 18): pc = Array(13, 16, 19, 22): dc = Array(7, 10, 13)
    End If
    For Each c In tc: t.Range(t.Cells(4, CLng(c)), t.Cells(last, CLng(c))).ClearContents: Next c
    For Each c In pc: p.Range(p.Cells(4, CLng(c)), p.Cells(last, CLng(c))).ClearContents: Next c
    For Each c In dc: d.Range(d.Cells(4, CLng(c)), d.Cells(6, CLng(c))).ClearContents: Next c
    AddLog "Skipped", life & " steady state", "No inputs for this life. Output values cleared; headings retained."
End Sub

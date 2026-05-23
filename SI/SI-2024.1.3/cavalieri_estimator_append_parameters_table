Option Explicit

Sub cavalieri_estimator_append_parameters_table()

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    Dim dstWs As Worksheet
    Dim topLeft As Range
    Dim headerRow As Long
    Dim startCol As Long
    Dim firstDataRow As Long

    Dim fd As FileDialog
    Dim filePath As Variant
    Dim srcWb As Workbook
    Dim wsSummary As Worksheet
    Dim wsParams As Worksheet

    Dim nextRow As Long
    Dim lastSummaryRow As Long
    Dim r As Long

    Dim excelFileNameValue As String
    Dim missingValue As Variant
    Dim dataFileValue As Variant
    Dim markerValue As Variant
    Dim shapeFactorValue As Variant
    Dim adjustedVolumeValue As Variant

    ' Destination = active sheet
    Set dstWs = ActiveSheet

    ' IMPORTANT:
    ' Select the existing "Excel File Name" header cell before running.
    Set topLeft = ActiveCell
    headerRow = topLeft.Row
    startCol = topLeft.Column
    firstDataRow = headerRow + 1

    ' Find next empty row below the selected Excel File Name column
    nextRow = NextEmptyRowFrom(dstWs, startCol, firstDataRow)

    Set fd = Application.FileDialog(msoFileDialogOpen)

    With fd
        .AllowMultiSelect = True
        .Filters.Clear
        .Filters.Add "Excel files", "*.xls*"
        .Title = "Select Cavalieri Estimator output files"

        If .Show <> -1 Then GoTo SafeExit

        For Each filePath In .SelectedItems

            Set srcWb = Nothing
            Set wsSummary = Nothing
            Set wsParams = Nothing

            On Error Resume Next
            Set srcWb = Workbooks.Open(CStr(filePath), ReadOnly:=True)
            On Error GoTo 0

            If srcWb Is Nothing Then
                MsgBox "Could not open file: " & CStr(filePath), vbExclamation
                GoTo NextFile
            End If

            excelFileNameValue = srcWb.Name

            On Error Resume Next
            Set wsSummary = srcWb.Worksheets("Summary")
            Set wsParams = srcWb.Worksheets("Parameters")
            On Error GoTo 0

            If wsSummary Is Nothing Then
                MsgBox "The file does not have a 'Summary' sheet: " & CStr(filePath), vbExclamation
                srcWb.Close SaveChanges:=False
                GoTo NextFile
            End If

            lastSummaryRow = wsSummary.Cells(wsSummary.Rows.Count, 1).End(xlUp).Row

            For r = 2 To lastSummaryRow

                If Application.WorksheetFunction.CountA(wsSummary.Rows(r)) > 0 Then

                    dataFileValue = GetByHeader(wsSummary, "Data File", r, srcWb.Name)
                    markerValue = GetByHeader(wsSummary, "Marker", r, "")

                    ' Missing sections can appear as "Missing Sections" or "Missing Section"
                    missingValue = GetByHeader(wsSummary, "Missing Sections", r, "")
                    
                    If IsEmpty(missingValue) Or Trim(CStr(missingValue)) = "" Then
                        missingValue = GetByHeader(wsSummary, "Missing Section", r, 0)
                    End If

                    If IsEmpty(missingValue) Or Trim(CStr(missingValue)) = "" Then
                        missingValue = 0
                    End If

                    missingValue = CleanNumber(missingValue)

                    ' Adjusted volume:
                    ' If there are no missing sections, report "none"
                    If IsNumeric(missingValue) And CDbl(missingValue) = 0 Then
                        adjustedVolumeValue = "none"
                    Else
                        adjustedVolumeValue = GetAdjustedVolume(wsSummary, r)

                        If IsEmpty(adjustedVolumeValue) Or Trim(CStr(adjustedVolumeValue)) = "" Then
                            adjustedVolumeValue = "not reported"
                        End If
                    End If

                    ' Shape Factor may be in Summary or Parameters
                    shapeFactorValue = GetByHeader(wsSummary, "Shape Factor", r, "")

                    If IsEmpty(shapeFactorValue) Or Trim(CStr(shapeFactorValue)) = "" Then
                        If Not wsParams Is Nothing Then
                            shapeFactorValue = GetParameterValue(wsParams, "Shape Factor", "")
                        End If
                    End If

                    ' Avoid duplicate rows based on Excel File Name + Data File + Structure
                    If Not RowAlreadyExists(dstWs, startCol, firstDataRow, _
                                            CStr(excelFileNameValue), _
                                            CStr(dataFileValue), _
                                            CStr(markerValue)) Then

                        With dstWs
                            .Cells(nextRow, startCol + 0).Value = excelFileNameValue
                            .Cells(nextRow, startCol + 1).Value = dataFileValue
                            .Cells(nextRow, startCol + 2).Value = markerValue

                            .Cells(nextRow, startCol + 3).Value = CleanNumber(GetByHeader(wsSummary, "Section Cut Thickness", r, ""))
                            .Cells(nextRow, startCol + 4).Value = CleanNumber(GetByHeader(wsSummary, "Section Evaluation Interval", r, ""))
                            .Cells(nextRow, startCol + 5).Value = CleanNumber(GetByHeader(wsSummary, "Grid Size", r, ""))

                            If Not wsParams Is Nothing Then
                                .Cells(nextRow, startCol + 6).Value = CleanNumber(GetParameterValue(wsParams, "Associated Area", ""))
                                .Cells(nextRow, startCol + 7).Value = CleanNumber(GetParameterValue(wsParams, "Associated Volume", ""))
                            Else
                                .Cells(nextRow, startCol + 6).Value = ""
                                .Cells(nextRow, startCol + 7).Value = ""
                            End If

                            .Cells(nextRow, startCol + 8).Value = CleanNumber(GetByHeader(wsSummary, "Sections", r, ""))
                            .Cells(nextRow, startCol + 9).Value = CleanNumber(GetByHeader(wsSummary, "Count", r, ""))
                            .Cells(nextRow, startCol + 10).Value = missingValue

                            .Cells(nextRow, startCol + 11).Value = CleanNumber(GetByHeader(wsSummary, "Estimated Volume", r, ""))
                            .Cells(nextRow, startCol + 12).Value = CleanNumber(adjustedVolumeValue)

                            .Cells(nextRow, startCol + 13).Value = CleanNumber(GetByHeader(wsSummary, "Coefficient of Error (Gundersen), m=1", r, ""))
                            .Cells(nextRow, startCol + 14).Value = CleanNumber(shapeFactorValue)
                        End With

                        nextRow = nextRow + 1

                    End If

                End If

            Next r

            srcWb.Close SaveChanges:=False

NextFile:
            Set wsSummary = Nothing
            Set wsParams = Nothing
            Set srcWb = Nothing

        Next filePath

    End With

    ApplyCavalieriNumberFormatting dstWs, startCol

    MsgBox "New Cavalieri samples were appended below the existing table.", vbInformation

SafeExit:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

End Sub


Private Sub ApplyCavalieriNumberFormatting(ws As Worksheet, startCol As Long)

    With ws
        ' Text columns
        .Columns(startCol + 0).NumberFormat = "@"
        .Columns(startCol + 1).NumberFormat = "@"
        .Columns(startCol + 2).NumberFormat = "@"

        ' Integer-like columns
        .Columns(startCol + 3).NumberFormat = "0"   ' Section cut thickness
        .Columns(startCol + 4).NumberFormat = "0"   ' Section evaluation interval
        .Columns(startCol + 5).NumberFormat = "0"   ' Grid size
        .Columns(startCol + 8).NumberFormat = "0"   ' No. sections analyzed
        .Columns(startCol + 9).NumberFormat = "0"   ' Total points counted
        .Columns(startCol + 10).NumberFormat = "0"  ' Missing sections

        ' Area per point
        .Columns(startCol + 6).NumberFormat = "#,##0"

        ' Associated volume
        .Columns(startCol + 7).NumberFormat = "0.000000"

        ' Estimated volume
        .Columns(startCol + 11).NumberFormat = "0.000"

        ' Estimated volume adjusted for missing sections
        .Columns(startCol + 12).NumberFormat = "General"

        ' CE Gundersen m=1
        .Columns(startCol + 13).NumberFormat = "0.000"

        ' Shape Factor
        .Columns(startCol + 14).NumberFormat = "0.000"

        .Range(.Cells(1, startCol), .Cells(1, startCol + 14)).EntireColumn.AutoFit
    End With

End Sub


Private Function GetAdjustedVolume(ws As Worksheet, dataRow As Long) As Variant

    Dim colNum As Long
    Dim val As Variant

    ' Best match: header contains volume + adjusted + missing + section
    colNum = FindHeaderColumnAll(ws, Array("volume", "adjusted", "missing", "section"))

    If colNum > 0 Then
        val = ws.Cells(dataRow, colNum).Value
        If Not IsEmpty(val) And Trim(CStr(val)) <> "" Then
            GetAdjustedVolume = val
            Exit Function
        End If
    End If

    ' Fallback: header contains mm3 + adjusted + missing + section
    colNum = FindHeaderColumnAll(ws, Array("mm3", "adjusted", "missing", "section"))

    If colNum > 0 Then
        val = ws.Cells(dataRow, colNum).Value
        If Not IsEmpty(val) And Trim(CStr(val)) <> "" Then
            GetAdjustedVolume = val
            Exit Function
        End If
    End If

    ' Fallback: header contains mm + adjusted + missing + section
    colNum = FindHeaderColumnAll(ws, Array("mm", "adjusted", "missing", "section"))

    If colNum > 0 Then
        val = ws.Cells(dataRow, colNum).Value
        If Not IsEmpty(val) And Trim(CStr(val)) <> "" Then
            GetAdjustedVolume = val
            Exit Function
        End If
    End If

    GetAdjustedVolume = ""

End Function


Private Function FindHeaderColumnAll(ws As Worksheet, requiredTerms As Variant) As Long

    Dim lastCol As Long
    Dim c As Long
    Dim rr As Long
    Dim i As Long
    Dim headerText As String
    Dim term As String
    Dim allFound As Boolean

    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    For c = 1 To lastCol

        headerText = ""

        ' Look across the first few rows in case the header is split across rows
        For rr = 1 To 5
            headerText = headerText & " " & NormalizeText(ws.Cells(rr, c).Value)
        Next rr

        allFound = True

        For i = LBound(requiredTerms) To UBound(requiredTerms)
            term = NormalizeText(requiredTerms(i))

            If InStr(1, headerText, term, vbTextCompare) = 0 Then
                allFound = False
                Exit For
            End If
        Next i

        If allFound Then
            FindHeaderColumnAll = c
            Exit Function
        End If

    Next c

    FindHeaderColumnAll = 0

End Function


Private Function NextEmptyRowFrom(ws As Worksheet, startCol As Long, firstDataRow As Long) As Long

    Dim lastRow As Long

    lastRow = ws.Cells(ws.Rows.Count, startCol).End(xlUp).Row

    If lastRow < firstDataRow Then
        NextEmptyRowFrom = firstDataRow
    Else
        NextEmptyRowFrom = lastRow + 1
    End If

End Function


Private Function RowAlreadyExists(ws As Worksheet, startCol As Long, firstDataRow As Long, _
                                  excelFileName As String, dataFile As String, structureName As String) As Boolean

    Dim lastRow As Long
    Dim r As Long

    lastRow = ws.Cells(ws.Rows.Count, startCol).End(xlUp).Row

    If lastRow < firstDataRow Then
        RowAlreadyExists = False
        Exit Function
    End If

    For r = firstDataRow To lastRow
        If Trim(CStr(ws.Cells(r, startCol + 0).Value)) = Trim(excelFileName) _
           And Trim(CStr(ws.Cells(r, startCol + 1).Value)) = Trim(dataFile) _
           And Trim(CStr(ws.Cells(r, startCol + 2).Value)) = Trim(structureName) Then
            RowAlreadyExists = True
            Exit Function
        End If
    Next r

    RowAlreadyExists = False

End Function


Private Function GetByHeader(ws As Worksheet, headerText As String, dataRow As Long, Optional defaultValue As Variant = "") As Variant

    Dim colNum As Long
    colNum = FindHeaderColumn(ws, headerText)

    If colNum = 0 Then
        GetByHeader = defaultValue
    Else
        GetByHeader = ws.Cells(dataRow, colNum).Value
    End If

End Function


Private Function FindHeaderColumn(ws As Worksheet, headerText As String) As Long

    Dim lastCol As Long
    Dim c As Long
    Dim target As String
    Dim currentText As String

    target = NormalizeText(headerText)
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    ' Exact match
    For c = 1 To lastCol
        currentText = NormalizeText(ws.Cells(1, c).Value)
        If currentText = target Then
            FindHeaderColumn = c
            Exit Function
        End If
    Next c

    ' Starts with target
    For c = 1 To lastCol
        currentText = NormalizeText(ws.Cells(1, c).Value)
        If Left(currentText, Len(target)) = target Then
            FindHeaderColumn = c
            Exit Function
        End If
    Next c

    ' Contains target
    For c = 1 To lastCol
        currentText = NormalizeText(ws.Cells(1, c).Value)
        If InStr(1, currentText, target, vbTextCompare) > 0 Then
            FindHeaderColumn = c
            Exit Function
        End If
    Next c

    FindHeaderColumn = 0

End Function


Private Function GetParameterValue(ws As Worksheet, parameterName As String, Optional defaultValue As Variant = "") As Variant

    Dim cell As Range
    Dim target As String
    Dim currentText As String

    target = NormalizeText(parameterName)

    For Each cell In ws.UsedRange.Cells
        currentText = NormalizeText(cell.Value)

        If currentText = target Or InStr(1, currentText, target, vbTextCompare) > 0 Then
            GetParameterValue = cell.Offset(0, 1).Value
            Exit Function
        End If
    Next cell

    GetParameterValue = defaultValue

End Function


Private Function NormalizeText(valueIn As Variant) As String

    Dim txt As String

    txt = CStr(valueIn)
    txt = Replace(txt, vbCr, " ")
    txt = Replace(txt, vbLf, " ")
    txt = Replace(txt, Chr(160), " ")

    ' Normalize symbols and superscripts
    txt = Replace(txt, "³", "3")
    txt = Replace(txt, "²", "2")
    txt = Replace(txt, "µ", "u")
    txt = Replace(txt, "μ", "u")

    txt = Application.WorksheetFunction.Trim(txt)
    txt = LCase(txt)

    NormalizeText = txt

End Function


Private Function CleanNumber(valueIn As Variant) As Variant

    Dim txt As String

    If IsEmpty(valueIn) Or Trim(CStr(valueIn)) = "" Then
        CleanNumber = ""
        Exit Function
    End If

    If IsNumeric(valueIn) Then
        CleanNumber = CDbl(valueIn)
        Exit Function
    End If

    txt = Trim(CStr(valueIn))
    txt = Replace(txt, Chr(160), "")
    txt = Replace(txt, " ", "")

    ' Remove thousands separator from exported values such as 2,500
    txt = Replace(txt, ",", "")

    ' Convert values such as "20." to "20"
    If Right(txt, 1) = "." Then
        txt = Left(txt, Len(txt) - 1)
    End If

    If IsNumeric(txt) Then
        CleanNumber = CDbl(txt)
    Else
        CleanNumber = valueIn
    End If

End Function

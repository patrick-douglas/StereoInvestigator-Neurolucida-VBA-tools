Sub cavalieri_estimator_data_summary()
Application.ScreenUpdating = False
'-----------------------------------
Dim xFd As FileDialog
Dim xFdItem As Variant
Dim xFileName As String
Dim xFB As String
With Application.FileDialog(msoFileDialogOpen)
.AllowMultiSelect = True
.Filters.Clear
.Filters.Add "excel", "*.xls*"
.Show
If .SelectedItems.Count < 1 Then Exit Sub
For lngCount = 1 To .SelectedItems.Count
xFileName = .SelectedItems(lngCount)
If xFileName <> "" Then
With Workbooks.Open(Filename:=xFileName)
sFileName = Split(xFileName, "\")(UBound(Split(xFileName, "\")))
'-----------------------------------
'NAME OF EXCEL WORKBOOK
    Windows(sFileName).Activate
    Sheets("Summary").Select
    Range("A1").Select
    ActiveCell.Offset(1, 26).Range("A1").Select
    ActiveCell.FormulaR1C1 = sFileName
    Selection.Copy
    Windows("Stereology-Summary-Table.xlsm").Activate
    ActiveCell.Offset(0, 0).Range("A1").Select
    ActiveCell.Offset(1, 0).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

'Number of missing sections
Windows(sFileName).Activate
Sheets("Summary").Select
Range("A1").Select
ActiveCell.Offset(1, 14).Range("A1").Select

Dim missingSectionValue As Variant
missingSectionValue = ActiveCell.Value

Windows("Stereology-Summary-Table.xlsm").Activate
Sheets("DataSummary-CavalieriEstimator").Select ' Replace "Sheet1" with the correct sheet name
'Range("A1").Select ' Replace "B2" with the desired destination cell
'    ActiveCell.Offset(0, 0).Range("A1").Select
    ActiveCell.Offset(0, 1).Range("A1").Select

If IsEmpty(missingSectionValue) Then
    ActiveCell.Value = "none"
Else
    ActiveCell.Value = missingSectionValue
End If

Range("B:B").NumberFormat = "0"
    
'Estimated Volume (mm³)
    Windows(sFileName).Activate
    Sheets("Summary").Select
    Range("A1").Select
    ActiveCell.Offset(1, 3).Range("A1").Select
    Selection.Copy
    Windows("Stereology-Summary-Table.xlsm").Activate
    ActiveCell.Offset(0, 0).Range("A1").Select
    ActiveCell.Offset(0, 1).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("C:C").NumberFormat = "0.000"
        
'Estimated Volume (mm³) adjusted with missing section counts
Windows(sFileName).Activate
Sheets("Summary").Select
Range("A1").Select
ActiveCell.Offset(1, 17).Range("A1").Select

Dim missingSectionValue2 As Variant
missingSectionValue2 = ActiveCell.Value

Windows("Stereology-Summary-Table.xlsm").Activate
Sheets("DataSummary-CavalieriEstimator").Select
    ActiveCell.Offset(0, 1).Range("A1").Select

If IsEmpty(missingSectionValue2) Then
    ActiveCell.Value = "none"
Else
    ActiveCell.Value = missingSectionValue2
End If

Range("B:B").NumberFormat = "0"
                   
'Coefficient of Error (Gundersen), m=0
    Windows(sFileName).Activate
    Sheets("Summary").Select
    Range("A1").Select
    ActiveCell.Offset(1, 5).Range("A1").Select
    Selection.Copy
    Windows("Stereology-Summary-Table.xlsm").Activate
    ActiveCell.Offset(0, 0).Range("A1").Select
    ActiveCell.Offset(0, 1).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("E:E").NumberFormat = "0.000"

'Coefficient of Error (Gundersen), m=1
    Windows(sFileName).Activate
    Sheets("Summary").Select
    Range("A1").Select
    ActiveCell.Offset(1, 6).Range("A1").Select
    Selection.Copy
    Windows("Stereology-Summary-Table.xlsm").Activate
    ActiveCell.Offset(0, 0).Range("A1").Select
    ActiveCell.Offset(0, 1).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("F:F").NumberFormat = "0.000"
                                  
'Stereology-Summary-Table.xlsm
    ActiveCell.Offset(0, -5).Range("A1").Select
    Windows("Stereology-Summary-Table.xlsm").Activate
    ActiveSheet.Cells.Columns("A:D").AutoFit
'    Range("B:B").NumberFormat = "#,##0.00"
    Workbooks(sFileName).Close SaveChanges:=False
Application.ScreenUpdating = False

End With
End If
Next lngCount
End With
End Sub
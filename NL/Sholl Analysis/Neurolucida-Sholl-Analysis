Sub neurolucida_summary_sholl_analysis()
Application.ScreenUpdating = False
'-----------------------------------
Dim xFd As FileDialog
Dim xFdItem As Variant
Dim xFileName As String
Dim xFB As String
Dim UserRadius As Variant
With Application.FileDialog(msoFileDialogOpen)
.AllowMultiSelect = True
.Filters.Clear
.Filters.Add "excel", "*.xls*"
.Show
'Radius selection
UserRadius = InputBox("Select a radius to extract (Only pairs e.g. 2,4,6,8,10 e etc):")
If .SelectedItems.Count < 1 Then Exit Sub
For lngCount = 1 To .SelectedItems.Count
xFileName = .SelectedItems(lngCount)
If xFileName <> "" Then
With Workbooks.Open(Filename:=xFileName)
sFileName = Split(xFileName, "\")(UBound(Split(xFileName, "\")))

'UserRadius to RealRadius  conversion
    RealRadius = UserRadius / 2

'-----------------------------------
'NAME OF EXCEL WORKBOOK
    Windows(sFileName).Activate
    Sheets("Sholl - Dendrites").Select
    Range("A1").Select
    ActiveCell.Offset(0, 17).Range("A1").Select
    ActiveCell.FormulaR1C1 = sFileName
    Selection.Copy
    Windows("Neurolucida-Sholl-Analysis.xlsm").Activate
    ActiveCell.Offset(0, 0).Range("A1").Select
    ActiveCell.Offset(1, 0).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
'Radius
    Windows(sFileName).Activate
    Sheets("Sholl - Dendrites").Select
    Range("A1").Select
    ActiveCell.Offset(RealRadius, 0).Range("A1").Select
    Selection.Copy
    Windows("Neurolucida-Sholl-Analysis.xlsm").Activate
    ActiveCell.Offset(0, 0).Range("A1").Select
    ActiveCell.Offset(0, 1).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
'Intersections
    Windows(sFileName).Activate
    Sheets("Sholl - Dendrites").Select
    Range("A1").Select
    ActiveCell.Offset(RealRadius, 1).Range("A1").Select
    Selection.Copy
    Windows("Neurolucida-Sholl-Analysis.xlsm").Activate
    ActiveCell.Offset(0, 0).Range("A1").Select
    ActiveCell.Offset(0, 1).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
'Length(µm)
    Windows(sFileName).Activate
    Sheets("Sholl - Dendrites").Select
    Range("A1").Select
    ActiveCell.Offset(RealRadius, 2).Range("A1").Select
    Selection.Copy
    Windows("Neurolucida-Sholl-Analysis.xlsm").Activate
    ActiveCell.Offset(0, 0).Range("A1").Select
    ActiveCell.Offset(0, 1).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
'SArea(µm²)
    Windows(sFileName).Activate
    Sheets("Sholl - Dendrites").Select
    Range("A1").Select
    ActiveCell.Offset(RealRadius, 3).Range("A1").Select
    Selection.Copy
    Windows("Neurolucida-Sholl-Analysis.xlsm").Activate
    ActiveCell.Offset(0, 0).Range("A1").Select
    ActiveCell.Offset(0, 1).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
'Volume(µm³)
    Windows(sFileName).Activate
    Sheets("Sholl - Dendrites").Select
    Range("A1").Select
    ActiveCell.Offset(RealRadius, 4).Range("A1").Select
    Selection.Copy
    Windows("Neurolucida-Sholl-Analysis.xlsm").Activate
    ActiveCell.Offset(0, 0).Range("A1").Select
    ActiveCell.Offset(0, 1).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
'Avg. Diameter
    Windows(sFileName).Activate
    Sheets("Sholl - Dendrites").Select
    Range("A1").Select
    ActiveCell.Offset(RealRadius, 5).Range("A1").Select
    Selection.Copy
    Windows("Neurolucida-Sholl-Analysis.xlsm").Activate
    ActiveCell.Offset(0, 0).Range("A1").Select
    ActiveCell.Offset(0, 1).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
'Nodes
    Windows(sFileName).Activate
    Sheets("Sholl - Dendrites").Select
    Range("A1").Select
    ActiveCell.Offset(RealRadius, 6).Range("A1").Select
    Selection.Copy
    Windows("Neurolucida-Sholl-Analysis.xlsm").Activate
    ActiveCell.Offset(0, 0).Range("A1").Select
    ActiveCell.Offset(0, 1).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
                
'Endings
    Windows(sFileName).Activate
    Sheets("Sholl - Dendrites").Select
    Range("A1").Select
    ActiveCell.Offset(RealRadius, 7).Range("A1").Select
    Selection.Copy
    Windows("Neurolucida-Sholl-Analysis.xlsm").Activate
    ActiveCell.Offset(0, 0).Range("A1").Select
    ActiveCell.Offset(0, 1).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

'End of workbook processing
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    ActiveCell.Offset(0, -8).Range("A1").Select
    Windows("Neurolucida-Sholl-Analysis.xlsm").Activate
    ActiveSheet.Cells.Columns("B:B").AutoFit
Workbooks(sFileName).Close SaveChanges:=False
Application.ScreenUpdating = False

End With
End If
Next lngCount
End With
End Sub

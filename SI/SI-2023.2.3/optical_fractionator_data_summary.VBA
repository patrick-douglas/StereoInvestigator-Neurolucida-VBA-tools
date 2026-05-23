Sub optical_fractionator_data_summary()
Application.ScreenUpdating = True
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
'your code

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
'Region
    Windows(sFileName).Activate
    Sheets("Summary").Select
    Range("A1").Select
    ActiveCell.Offset(1, 4).Range("A1").Select
    Selection.Copy
    Windows("Stereology-Summary-Table.xlsm").Activate
    ActiveCell.Offset(0, 0).Range("A1").Select
    ActiveCell.Offset(0, 1).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
'Marker
    Windows(sFileName).Activate
    Sheets("Summary").Select
    Range("A1").Select
    ActiveCell.Offset(1, 1).Range("A1").Select
    Selection.Copy
    Windows("Stereology-Summary-Table.xlsm").Activate
    ActiveCell.Offset(0, 0).Range("A1").Select
    ActiveCell.Offset(0, 1).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

'--------------------------------------------------------

'Estimated Population using Number Weighted Section Thickness
Windows(sFileName).Activate
Sheets("Summary").Select
Range("A1").Select

Dim firstCell As Range
Dim firstValue As Variant
Dim secondValue As Variant

'Procurar a célula que contém "Estimated Population using Number Weighted Section Thickness"
Set firstCell = Cells.Find(What:="Estimated Population using Number Weighted Section Thickness", LookIn:=xlValues, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=False)

If Not firstCell Is Nothing Then
    ' Obter o valor abaixo da célula encontrada
    firstValue = firstCell.Offset(1, 0).Value

    ' Verificar se a célula ao lado direito contém "Estimate adjusted by Missing Section Fraction"
    If firstCell.Offset(0, 1).Value = "Estimate adjusted by Missing Section Fraction" Then
        ' Obter o valor ao lado direito do valor encontrado
        secondValue = firstCell.Offset(1, 1).Value
    Else
        secondValue = "No Missing Sections"
    End If

    ' Ativar a planilha "Stereology-Summary-Table.xlsm"
    Windows("Stereology-Summary-Table.xlsm").Activate
    ' Colar o primeiro valor
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.Offset(0, 0).Value = firstValue
    ' Colar o segundo valor na célula à direita
    ActiveCell.Offset(0, 1).Value = secondValue
    ActiveCell.Offset(0, 1).Range("A1").Select
End If

'Missing section number
Windows(sFileName).Activate
Sheets("Summary").Select
Range("A1").Select

Dim cellEstimate As Range
Dim valueEstimate As Variant
Dim valueMissingSections As Variant

'Procurar a célula que contém "Estimated Population using Number Weighted Section Thickness"
Set cellEstimate = Cells.Find(What:="Estimated Population using Number Weighted Section Thickness", LookIn:=xlValues, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=False)

If Not cellEstimate Is Nothing Then
    ' Obter o valor abaixo da célula encontrada
    valueEstimate = cellEstimate.Offset(1, 0).Value

    ' Verificar se a célula ao lado direito contém "Estimate adjusted by Missing Section Fraction"
    If cellEstimate.Offset(0, 1).Value = "Estimate adjusted by Missing Section Fraction" Then
        ' Obter o valor ao lado direito do valor encontrado
        valueMissingSections = Range("Q2").Offset(0, 0).Value
    Else
        ' Verificar se há "Missing Sections" diretamente na célula Q2
        If Range("Q1").Value = "Missing Sections" Then
            valueMissingSections = Range("Q2").Offset(0, 0).Value
        Else
            valueMissingSections = "none"  ' Se não houver "Missing Sections", atribuir "none"
        End If
    End If

    ' Ativar a planilha "Stereology-Summary-Table.xlsm"
    Windows("Stereology-Summary-Table.xlsm").Activate
    ' Colar o primeiro valor
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.Offset(0, 0).Value = valueMissingSections
End If
        
'Estimated CE (Scheaffer)
    Windows(sFileName).Activate
    Sheets("Coefficient of Error").Select
    Range("A1").Select
    ActiveCell.Offset(1, 7).Range("A1").Select
    Selection.Copy
    Windows("Stereology-Summary-Table.xlsm").Activate
    ActiveCell.Offset(0, 0).Range("A1").Select
    ActiveCell.Offset(0, 1).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("C:C").NumberFormat = "0.0000"

'Measured Defined Mounted Thickness
    Windows(sFileName).Activate
    Sheets("Summary").Select
    Range("A1").Select
    ActiveCell.Offset(1, 7).Range("A1").Select
    Selection.Copy
    Windows("Stereology-Summary-Table.xlsm").Activate
    ActiveCell.Offset(0, 0).Range("A1").Select
    ActiveCell.Offset(0, 1).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("D:D").NumberFormat = "0.0"

                                  
'Stereology-Summary-Table.xlsm
    ActiveCell.Offset(0, -7).Range("A1").Select
    Windows("Stereology-Summary-Table.xlsm").Activate
    ActiveSheet.Cells.Columns("A:D").AutoFit
    Range("B:B").NumberFormat = "#,##0.00"
    Workbooks(sFileName).Close SaveChanges:=False
    
Columns("A:Z").NumberFormat = "0.0000"
Columns("A:B").AutoFit
Columns("E:E").AutoFit
Columns("D:D").NumberFormat = "#,##0.0"
Columns("E:E").NumberFormat = "#,##0.0"
Columns("A:Z").Font.Bold = False
Columns("F:F").NumberFormat = "0"
Application.ScreenUpdating = True
Columns("A:A").HorizontalAlignment = xlLeft
With Columns("B:Z")
    .HorizontalAlignment = xlCenter
End With


End With
End If
Next lngCount
End With
End Sub
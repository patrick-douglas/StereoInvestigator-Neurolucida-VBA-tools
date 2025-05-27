Sub optical_fractionator_data_summary()
Application.ScreenUpdating = False
'-----------------------------------
' Selection of xlsx file
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
'------------------------------------------------


'Caixa de Dialogo para o marcador
Dim markersDict As Object
Dim cell As Range
Dim marker As Variant
Dim prompt As String
Dim userSelection As Variant
Dim isValid As Boolean
Dim selectedMarker As Long
Dim i As Long

Set markersDict = CreateObject("Scripting.Dictionary")

With Sheets("Summary")
    Dim lastRow As Long
    lastRow = .Cells(.Rows.Count, "B").End(xlUp).Row
    For Each cell In .Range("B2:B" & lastRow)
        marker = Trim(cell.Value)
        If marker <> "" Then
            If Not markersDict.exists(marker) Then
                markersDict.Add marker, markersDict.Count + 1
            End If
        End If
    Next cell
End With

'Monta prompt numerado
prompt = "Available markers:" & vbCrLf
For Each marker In markersDict.Keys
    prompt = prompt & markersDict(marker) & " - " & marker & vbCrLf
Next
prompt = prompt & vbCrLf & "Please type the number of the marker to select:"

' Loop para receber número válido
Do
    userSelection = Application.InputBox(prompt, "Select Marker", Type:=1) ' Tipo 1 = número
    If userSelection = False Then ' Cancelou
        MsgBox "Selection cancelled.", vbInformation
        Exit Sub
    End If
    
    isValid = False
    If IsNumeric(userSelection) Then
        selectedMarker = CLng(userSelection)
        If selectedMarker >= 1 And selectedMarker <= markersDict.Count Then
            isValid = True
        End If
    End If
    
    If Not isValid Then
        MsgBox "Invalid selection. Please enter a number between 1 and " & markersDict.Count, vbExclamation
    End If
Loop While Not isValid

'------------------------------------------------


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
    ActiveCell.Offset(selectedMarker, 4).Range("A1").Select
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
    ActiveCell.Offset(selectedMarker, 1).Range("A1").Select
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

'Procurar a c_lula que cont_m "Estimated Population using Number Weighted Section Thickness"
Set firstCell = Cells.Find(What:="Estimated Population using Number Weighted Section Thickness", LookIn:=xlValues, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=False)

If Not firstCell Is Nothing Then
    ' Obter o valor abaixo da clula encontrada
    firstValue = firstCell.Offset(selectedMarker, 0).Value

    ' Verificar se a c_lula ao lado direito cont_m "Estimate adjusted by Missing Section Fraction"
    If firstCell.Offset(0, 1).Value = "Estimate adjusted by Missing Section Fraction" Then
        ' Obter o valor ao lado direito do valor encontrado
        secondValue = firstCell.Offset(selectedMarker, 1).Value
    Else
        secondValue = "No Missing Sections"
    End If

    ' Ativar a planilha "Stereology-Summary-Table.xlsm"
    Windows("Stereology-Summary-Table.xlsm").Activate
    ' Colar o primeiro valor
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.Offset(0, 0).Value = firstValue
    ' Colar o segundo valor na c_lula ö direita
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

'Procurar a c_lula que cont_m "Estimated Population using Number Weighted Section Thickness"
Set cellEstimate = Cells.Find(What:="Estimated Population using Number Weighted Section Thickness", LookIn:=xlValues, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=False)

If Not cellEstimate Is Nothing Then
    ' Obter o valor abaixo da c_lula encontrada
    valueEstimate = cellEstimate.Offset(1, 0).Value

    ' Verificar se a c_lula ao lado direito cont_m "Estimate adjusted by Missing Section Fraction"
    If cellEstimate.Offset(0, 1).Value = "Estimate adjusted by Missing Section Fraction" Then
        ' Obter o valor ao lado direito do valor encontrado
        valueMissingSections = Range("Q2").Offset(0, 0).Value
    Else
        ' Verificar se hà "Missing Sections" diretamente na c_lula Q2
        If Range("Q1").Value = "Missing Sections" Then
            valueMissingSections = Range("Q2").Offset(0, 0).Value
        Else
            valueMissingSections = "none"  ' Se nÜo houver "Missing Sections", atribuir "none"
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
    ActiveCell.Offset(selectedMarker, 7).Range("A1").Select
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
    ActiveCell.Offset(selectedMarker, 7).Range("A1").Select
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
Application.ScreenUpdating = False
Columns("A:A").HorizontalAlignment = xlLeft
With Columns("B:Z")
    .HorizontalAlignment = xlCenter
End With


End With
End If
Next lngCount
End With
End Sub
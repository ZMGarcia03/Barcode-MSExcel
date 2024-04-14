Sub GenerateBarcode()
    Dim cellValue As String
    Dim selectedCell As Range

    ' Get the value from the selected cell
    Set selectedCell = Application.InputBox("Select a cell containing the text or link:", Type:=8)
    cellValue = selectedCell.Value

    ' Set the font to Code 128
    selectedCell.Font.Name = "Code 128"

    ' Set the value to the same value, which will display the barcode
    selectedCell.Value = cellValue
End Sub

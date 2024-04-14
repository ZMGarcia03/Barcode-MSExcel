# Barcode-MSExcel
This [VBA macro code](Barcode.vba) generates a barcode of the given text or link in a specified cell in MS Excel using the [`Code128`](https://github.com/ZMGarcia03/Barcode-MSExcel/releases/download/Font/code_128.zip) font. The macro prompts the user to select a cell containing the text or link. It then sets the font of the selected cell to `Code128` and sets the value to the same value, which will display the barcode.

## How to Use:

  1. Open Excel:
     - Open Microsoft Excel.
  2. Open VBA Editor:
     - Press `Alt` + `F11` to open the VBA Editor.
  3. Insert Macro:
     - In the VBA Editor, go to `Insert` > `Module` to insert a new module.
  4. Copy and Paste Code:
     - Copy the provided VBA macro code and paste it into the module.
  5. Run Macro:
     - Close the VBA Editor.
     - Select a cell containing the text or link where you want to generate the barcode.
     - Go to `Developer` > `Macros`, select `GenerateBarcode`, and click Run.
  6. View Barcode:
     - The selected cell will now display the barcode generated from the text or link.

> [!NOTE]
> This macro provides a simple way to generate barcodes directly in MS Excel. Ensure that the [`Code128`](https://github.com/ZMGarcia03/Barcode-MSExcel/releases/download/Font/code_128.zip) font is installed on your system for the barcode to display correctly.

### LICENSE
This porject is protected under [MIT License](LICENSE). :shipit:

Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True
Set xlVbscript = objExcel.WorkBooks.Open("C:\Excelfile.xlsx")
CONST xlPasteValues = -4163
CONST xlPasteFormulas = -4123


xlVbscript.Sheets(1).Rows(3).EntireRow.Copy'''''''''''''Copy Entire 3rd Row''''''''''''''
xlVbscript.Sheets(1).Range("A:A").Copy'''''''''''''Copy Entire "A" Column''''''''''''''
xlVbscript.Sheets(1).Range("E1").PasteSpecial xlPasteValues'''''''Paste Copied Values Into "E" Column''''


xlVbscript.Sheets(1).Range("A1:A6").Copy''''''''''''''''Copy "A1" To "A6"''''''''''''
xlVbscript.Sheets(1).Range("D1").PasteSpecial xlPasteFormulas''''''''Paste Copied Formulas Into "D1" '''''''''

xlVbscript.save
xlVbscript.Close 

objExcel.Quit
set objExcel=nothing
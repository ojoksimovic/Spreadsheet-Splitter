Attribute VB_Name = "Module1"
Sub SplitandSaveSheet()

'This code is to split the worksheet by units and into the Graduate Unit SGSDrive

Dim Splitcode As Range
Dim Path As String
Dim F2 As String
Dim Workbook_Name As String
Dim Worksheet_Name As String
Dim Masterdata_Range As Range
Dim Department_Range As Range
Dim NewSheet_Title As String
Dim Division As String
Dim Department_Field As String
Dim Template_Range As Range
'Dim User_Name As String

'User_Name = Application.InputBox("What is your UTORID? All lowercaps, please!", _
Type:=2)

Set Masterdata_Range = Application.InputBox("Select the range of the entire sheet to be copied (including the headers).", _
Type:=8)

ActiveWorkbook.Names.Add _
            Name:="Masterdata", _
            RefersTo:=Masterdata_Range
         
Set Template_Range = Application.InputBox("Select the range of the results data table only (including the one header).", _
Type:=8)

ActiveWorkbook.Names.Add _
            Name:="Template", _
            RefersTo:=Template_Range
                  
Set Department_Range = Application.InputBox("Select the entire range of units. Do not include headers. Repeated units are OK.", _
Type:=8)

Workbook_Name = ActiveWorkbook.Name
Worksheet_Name = ActiveSheet.Name

Department_Field = Department_Range.Column

NewSheet_Title = Application.InputBox("Enter the name of the worksheets to be created. The unit code will automatically be added to the end of the name.", _
Type:=2)

Sheets.Add(After:=Sheets(Worksheet_Name)).Name = "Departments_VBA"
Department_Range.Copy Worksheets("Departments_VBA").Range("A1")

Worksheets("Departments_VBA").Range(Range("A1"), Range("A1").End(xlDown)).RemoveDuplicates Columns:=1, Header:=xlNo
ActiveWorkbook.Names.Add _
            Name:="Splitcode", _
            RefersTo:=Worksheets("Departments_VBA").Range(Range("A1"), Range("A1").End(xlDown))

'Path = "C:\Users\" & User_Name & "\OneDrive - University of Toronto\SGS Drive - SGS Drive\Graduate Units\"
Path = "C:\Users\joksimov\OneDrive - University of Toronto\SGS Drive - SGS Drive\Graduate Units\"

For Each cell In Range("Splitcode")

Workbooks(Workbook_Name).Worksheets(Worksheet_Name).Copy After:=Worksheets(Sheets.Count)
ActiveSheet.Name = cell.Value

With ActiveWorkbook.Sheets(cell.Value).Range("Template")
.AutoFilter Field:=Department_Field, Criteria1:="<>" & cell.Value, Criteria2:="<>" & "ABC", Operator:=xlFilterValues
.Offset(1, 0).SpecialCells(xlCellTypeVisible).EntireRow.Delete
End With

ActiveSheet.AutoFilter.ShowAllData

Workbooks(Workbook_Name).Sheets(cell.Value).Copy Before:=Workbooks.Add.Sheets(1)

Application.ActiveWorkbook.SaveAs Filename:=Path & cell.Value & "\" & NewSheet_Title & " - " & cell.Value, _
        FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
ActiveWorkbook.Close

Application.DisplayAlerts = False
Application.Workbooks(Workbook_Name).Worksheets(cell.Value).Delete
Application.DisplayAlerts = True

Next cell

Application.Workbooks(Workbook_Name).Names("Splitcode").Delete
Application.Workbooks(Workbook_Name).Names("Masterdata").Delete
Application.Workbooks(Workbook_Name).Names("Template").Delete
Application.DisplayAlerts = False
Application.ActiveWorkbook.Worksheets("Departments_VBA").Delete
Application.DisplayAlerts = True
End Sub




# Calculate the sum of a range of cells in a spreadsheet document

This topic shows how to use the classes in the Open XML SDK 2.5 for Office to calculate the sum of a contiguous range of cells in a spreadsheet document programmatically.

The following assembly directives are required to compile the code in this topic.

```csharp
    using System.Collections.Generic;
    using System.Linq;
    using System.Text.RegularExpressions;
    using DocumentFormat.OpenXml;
    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.Spreadsheet;
```



## Get a SpreadsheetDocument Object

In the Open XML SDK, the **[SpreadsheetDocument](https://msdn.microsoft.com/library/office/documentformat.openxml.packaging.spreadsheetdocument.aspx)** class represents an Excel document package. To open and work with an Excel document, you create an instance of the **SpreadsheetDocument** class from the document. After you create the instance from the document, you can then obtain access to the main **[Workbook](https://msdn.microsoft.com/library/office/documentformat.openxml.spreadsheet.workbook.aspx)** part that contains the worksheets. The text in the document is represented in the package as XML using **SpreadsheetML** markup.

To create the class instance from the document that you call one of the **Open** methods. Several are provided, each with a different signature. The sample code in this topic uses the **[Open(String, Boolean)](https://msdn.microsoft.com/library/office/cc562356.aspx)** method with a signature that requires two parameters. The first  parameter takes a full path string that represents the document that you want to open. The second parameter is either **true** or **false** and represents whether you want the file to be opened for editing. Any changes that you make to the document will not be saved if this parameter is **false**.

The code that calls the **Open** method is shown in the following **using** statement.

```csharp
    // Open the document for editing.
    using (SpreadsheetDocument document = SpreadsheetDocument.Open(docName, true)) 
    {
        // Other code goes here.
    }
```



The **using** statement provides a recommended alternative to the typical .Open, .Save, .Close sequence. It ensures that the **Dispose** method (internal method used by the Open XML SDK to clean up resources) is automatically called when the closing brace is reached. The block that follows the **using** statement establishes a scope for the object that is created or named in the **using** statement, in this case **document**.

## Basic Structure of a SpreadsheetML Document

The basic document structure of a **SpreadsheetML** document consists of the **[Sheets](https://msdn.microsoft.com/library/office/documentformat.openxml.spreadsheet.sheets.aspx)** and **[Sheet](https://msdn.microsoft.com/library/office/documentformat.openxml.spreadsheet.sheet.aspx)** elements, which reference the worksheets in the workbook. A separate XML file is created for each worksheet. For example, the **SpreadsheetML** for a workbook that has two worksheets name MySheet1 and MySheet2 is
located in the Workbook.xml file and is shown in the following code example.

```xml
    <?xml version="1.0" encoding="UTF-8" standalone="yes" ?> 
    <workbook xmlns=https://schemas.openxmlformats.org/spreadsheetml/2006/main xmlns:r="https://schemas.openxmlformats.org/officeDocument/2006/relationships">
        <sheets>
            <sheet name="MySheet1" sheetId="1" r:id="rId1" /> 
            <sheet name="MySheet2" sheetId="2" r:id="rId2" /> 
        </sheets>
    </workbook>
```

The worksheet XML files contain one or more block level elements such as **[SheetData](https://msdn.microsoft.com/library/office/documentformat.openxml.spreadsheet.sheetdata.aspx)**. **sheetData** represents the cell table and contains one or more **[Row](https://msdn.microsoft.com/library/office/documentformat.openxml.spreadsheet.row.aspx)** elements. A **row** contains one or more **[Cell](https://msdn.microsoft.com/library/office/documentformat.openxml.spreadsheet.cell.aspx)** elements. Each cell contains a **[CellValue](https://msdn.microsoft.com/library/office/documentformat.openxml.spreadsheet.cellvalue.aspx)** element that represents the value of the cell. For example, the SpreadsheetML for the first worksheet in a workbook, that only has the value 100 in cell A1, is located in the Sheet1.xml file and is shown in the following code example.

```xml
    <?xml version="1.0" encoding="UTF-8" ?> 
    <worksheet xmlns="https://schemas.openxmlformats.org/spreadsheetml/2006/main">
        <sheetData>
            <row r="1">
                <c r="A1">
                    <v>100</v> 
                </c>
            </row>
        </sheetData>
    </worksheet>
```

Using the Open XML SDK 2.5, you can create document structure and content that uses strongly-typed classes that correspond to **SpreadsheetML** elements. You can find these classes in the **DocumentFormat.OpenXML.Spreadsheet** namespace. The following table lists the class names of the classes that correspond to the **workbook**, **sheets**, **sheet**, **worksheet**, and **sheetData** elements.

| SpreadsheetML Element | Open XML SDK 2.5 Class | Description |
|---|---|---|
| workbook | DocumentFormat.OpenXml.Spreadsheet.Workbook | The root element for the main document part. |
| sheets | DocumentFormat.OpenXml.Spreadsheet.Sheets | The container for the block level structures such as sheet, fileVersion, and others specified in the [ISO/IEC 29500](https://www.iso.org/standard/71691.html) specification. |
| sheet | DocumentFormat.OpenXml.Spreadsheet.Sheet | A sheet that points to a sheet definition file. |
| worksheet | DocumentFormat.OpenXml.Spreadsheet.Worksheet | A sheet definition file that contains the sheet data. |
| sheetData | DocumentFormat.OpenXml.Spreadsheet.SheetData | The cell table, grouped together by rows. |
| row | DocumentFormat.OpenXml.Spreadsheet.Row | A row in the cell table. |
| c | DocumentFormat.OpenXml.Spreadsheet.Cell | A cell in a row. |
| v | DocumentFormat.OpenXml.Spreadsheet.CellValue | The value of a cell. |

## How the Sample Code Works

The sample code starts by passing in to the method **CalculateSumOfCellRange** a parameter that represents the full path to the source **SpreadsheetML** file, a parameter that represents the name of the worksheet that contains the cells, a parameter that represents the name of the first cell in the contiguous range, a parameter that represent the name of the last cell in the contiguous range, and a parameter that represents the name of the cell where you want the result displayed.

The code then opens the file for editing as a **SpreadsheetDocument** document package for read/write access, the code gets the specified **Worksheet** object. It then gets the index of the row for the first and last cell in the contiguous range by calling the **GetRowIndex** method. It gets the name of the column for the first and last cell in the contiguous range by calling the **GetColumnName** method.

For each **Row** object within the contiguous range, the code iterates through each **Cell** object and determines if the column of the cell is within the contiguous
range by calling the **CompareColumn** method. If the cell is within the contiguous range, the code adds the value of the cell to the sum. Then it gets the **SharedStringTablePart** object if it exists. If it does not exist, it creates one using the **[AddNewPart](https://msdn.microsoft.com/library/office/documentformat.openxml.packaging.openxmlpartcontainer.addnewpart.aspx)** method. It inserts the result into the **SharedStringTablePart** object by calling the **InsertSharedStringItem** method.

The code inserts a new cell for the result into the worksheet by calling the **InsertCellInWorksheet** method and set the value of the cell. For more information, see [how to insert a cell in a spreadsheet](how-to-insert-text-into-a-cell-in-a-spreadsheet.md#how-the-sample-code-works), and then saves the worksheet.

```csharp
    // Given a document name, a worksheet name, the name of the first cell in the contiguous range, 
    // the name of the last cell in the contiguous range, and the name of the results cell, 
    // calculates the sum of the cells in the contiguous range and inserts the result into the results cell.
    // Note: All cells in the contiguous range must contain numbers.
    private static void CalculateSumOfCellRange(string docName, string worksheetName, string firstCellName, string lastCellName, string resultCell)
    {
        // Open the document for editing.
        using (SpreadsheetDocument document = SpreadsheetDocument.Open(docName, true))
        {
            IEnumerable<Sheet> sheets = document.WorkbookPart.Workbook.Descendants<Sheet>().Where(s => s.Name == worksheetName);
            if (sheets.Count() == 0)
            {
                // The specified worksheet does not exist.
                return; 
            }

            WorksheetPart worksheetPart = (WorksheetPart)document.WorkbookPart.GetPartById(sheets.First().Id);
            Worksheet worksheet = worksheetPart.Worksheet;

            // Get the row number and column name for the first and last cells in the range.
            uint firstRowNum = GetRowIndex(firstCellName);
            uint lastRowNum = GetRowIndex(lastCellName);
            string firstColumn = GetColumnName(firstCellName);
            string lastColumn = GetColumnName(lastCellName);

            double sum = 0;

            // Iterate through the cells within the range and add their values to the sum.
            foreach (Row row in worksheet.Descendants<Row>().Where(r => r.RowIndex.Value >= firstRowNum && r.RowIndex.Value <= lastRowNum))
            {
                foreach (Cell cell in row)
                {
                    string columnName = GetColumnName(cell.CellReference.Value);
                    if (CompareColumn(columnName, firstColumn) >= 0 && CompareColumn(columnName, lastColumn) <= 0)
                    {
                        sum += double.Parse(cell.CellValue.Text);
                    }
                }
            }

            // Get the SharedStringTablePart and add the result to it.
            // If the SharedStringPart does not exist, create a new one.
            SharedStringTablePart shareStringPart;
            if (document.WorkbookPart.GetPartsOfType<SharedStringTablePart>().Count() > 0)
            {
                shareStringPart = document.WorkbookPart.GetPartsOfType<SharedStringTablePart>().First();
            }
            else
            {
                shareStringPart = document.WorkbookPart.AddNewPart<SharedStringTablePart>();
            }

            // Insert the result into the SharedStringTablePart.
            int index = InsertSharedStringItem("Result:" + sum, shareStringPart);

            Cell result = InsertCellInWorksheet(GetColumnName(resultCell), GetRowIndex(resultCell), worksheetPart);

            // Set the value of the cell.
            result.CellValue = new CellValue(index.ToString());
            result.DataType = new EnumValue<CellValues>(CellValues.SharedString);

            worksheetPart.Worksheet.Save();
        }
    }
```


To get the row index the code passes a parameter that represents the name of the cell, and creates a new regular expression to match the row
index portion of the cell name. For more information about regular expressions, see [Regular Expression Language Elements](/dotnet/standard/base-types/regular-expression-language-quick-reference.md). It gets the row index by calling the **[Regex.Match](https://msdn2.microsoft.com/library/3zy662f6)** method, and then returns the row index.

```csharp
    // Given a cell name, parses the specified cell to get the row index.
    private static uint GetRowIndex(string cellName)
    {
        // Create a regular expression to match the row index portion the cell name.
        Regex regex = new Regex(@"\d+");
        Match match = regex.Match(cellName);

        return uint.Parse(match.Value);
    }
```



The code then gets the column name by passing a parameter that represents the name of the cell, and creates a new regular expression to match the column name portion of the cell name. This regular expression matches any combination of uppercase or lowercase letters. It gets the column name by calling the **[Regex.Match](/dotnet/api/system.text.regularexpressions.regex.match.md)** method, and then returns the column name.

```csharp
    // Given a cell name, parses the specified cell to get the column name.
    private static string GetColumnName(string cellName)
    {
        // Create a regular expression to match the column name portion of the cell name.
        Regex regex = new Regex("[A-Za-z]+");
        Match match = regex.Match(cellName);

        return match.Value;
    }
```



To compare two columns the code passes in two parameters that represent the columns to compare. If the first column is longer than the second column, it returns 1. If the second column is longer than the first column, it returns -1. Otherwise, it compares the values of the columns using the **[Compare](/dotnet/api/system.string.compare?view=net-6.0)** and returns the result.

```csharp
    // Given two columns, compares the columns.
    private static int CompareColumn(string column1, string column2)
    {
        if (column1.Length > column2.Length)
        {
            return 1;
        }
        else if (column1.Length < column2.Length)
        {
            return -1;
        }
        else
        {
            return string.Compare(column1, column2, true);
        }
    }
```



To insert a **SharedStringItem**, the code passes in a parameter that represents the text to insert into the cell and a parameter that represents the  **SharedStringTablePart** object for the spreadsheet. If the **ShareStringTablePart** object does not contain a **[SharedStringTable](https://msdn.microsoft.com/library/office/documentformat.openxml.spreadsheet.sharedstringtable.aspx)** object then it creates one. If the text already exists in the **ShareStringTable** object, then it returns the index for the **[SharedStringItem](/dotnet/api/documentformat.openxml.spreadsheet.sharedstringitem.md)** object that represents the text. If the text does not exist, create a new **SharedStringItem** object that represents the text. It then returns the index for the **SharedStringItem** object that represents the text.

```csharp
    // Given text and a SharedStringTablePart, creates a SharedStringItem with the specified text 
    // and inserts it into the SharedStringTablePart. If the item already exists, returns its index.
    private static int InsertSharedStringItem(string text, SharedStringTablePart shareStringPart)
    {
        // If the part does not contain a SharedStringTable, create it.
        if (shareStringPart.SharedStringTable == null)
        {
            shareStringPart.SharedStringTable = new SharedStringTable();
        }

        int i = 0;
        foreach (SharedStringItem item in shareStringPart.SharedStringTable.Elements<SharedStringItem>())
        {
            if (item.InnerText == text)
            {
                // The text already exists in the part. Return its index.
                return i;
            }

            i++;
        }

        // The text does not exist in the part. Create the SharedStringItem.
        shareStringPart.SharedStringTable.AppendChild(new SharedStringItem(new DocumentFormat.OpenXml.Spreadsheet.Text(text)));
        shareStringPart.SharedStringTable.Save();

        return i;
    }
```



The final step is to insert a cell into the worksheet. The code does that by passing in parameters that represent the name of the column and the number of the row of the cell, and a parameter that represents the worksheet that contains the cell. If the specified row does not exist, it creates the row and append it to the worksheet. If the specified column exists, it finds the cell that matches the row in that column and returns the cell. If the specified column does not exist, it creates the column and inserts it into the worksheet. It then determines where to insert the new cell in the column by iterating through the row elements to find the cell that comes directly after the specified row, in sequential order. It saves this row in the **refCell** variable. It inserts the new cell before the cell referenced by **refCell** using the **[InsertBefore](/dotnet/api/documentformat.openxml.openxmlcompositeelement.insertbefore.md)** method. It then returns the new **Cell** object.

```csharp
    // Given a column name, a row index, and a WorksheetPart, inserts a cell into the worksheet. 
    // If the cell already exists, returns it. 
    private static Cell InsertCellInWorksheet(string columnName, uint rowIndex, WorksheetPart worksheetPart)
    {
        Worksheet worksheet = worksheetPart.Worksheet;
        SheetData sheetData = worksheet.GetFirstChild<SheetData>();
        string cellReference = columnName + rowIndex;

        // If the worksheet does not contain a row with the specified row index, insert one.
        Row row;
        if (sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).Count() != 0)
        {
            row = sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).First();
        }
        else
        {
            row = new Row() { RowIndex = rowIndex };
            sheetData.Append(row);
        }

        // If there is not a cell with the specified column name, insert one.  
        if (row.Elements<Cell>().Where(c => c.CellReference.Value == columnName + rowIndex).Count() > 0)
        {
            return row.Elements<Cell>().Where(c => c.CellReference.Value == cellReference).First();
        }
        else
        {
            // Cells must be in sequential order according to CellReference. Determine where to insert the new cell.
            Cell refCell = null;
            foreach (Cell cell in row.Elements<Cell>())
            {
                if (string.Compare(cell.CellReference.Value, cellReference, true) > 0)
                {
                    refCell = cell;
                    break;
                }
            }

            Cell newCell = new Cell() { CellReference = cellReference };
            row.InsertBefore(newCell, refCell);

            worksheet.Save();
            return newCell;
        }
    }
```



## Sample Code

The following code sample calculates the sum of a contiguous range of cells in a spreadsheet document. The result is inserted into the **SharedStringTablePart** object and into the specified result cell. You can call the method CalculateSumOfCellRange by using the following example.

```csharp
    string docName = @"C:\Users\Public\Documents\Sheet1.xlsx";
    string worksheetName = "John";
    string firstCellName = "A1";
    string lastCellName = "A3";
    string resultCell = "A4";
    CalculateSumOfCellRange(docName, worksheetName, firstCellName, lastCellName, resultCell);
```



After running the program, you can inspect the file named "Sheet1.xlsx" to see the sum of the column in the worksheet named "John" in the specified cell.

The following is the complete sample code in both C\# and Visual Basic.

```csharp
    private static void CalculateSumOfCellRange(string docName, string worksheetName, string firstCellName, string lastCellName, string resultCell)
    {
        // Open the document for editing.
        using (SpreadsheetDocument document = SpreadsheetDocument.Open(docName, true))
        {
            IEnumerable<Sheet> sheets = document.WorkbookPart.Workbook.Descendants<Sheet>().Where(s => s.Name == worksheetName);
            if (sheets.Count() == 0)
            {
                // The specified worksheet does not exist.
                return;
            }

            WorksheetPart worksheetPart = (WorksheetPart)document.WorkbookPart.GetPartById(sheets.First().Id);
            Worksheet worksheet = worksheetPart.Worksheet;

            // Get the row number and column name for the first and last cells in the range.
            uint firstRowNum = GetRowIndex(firstCellName);
            uint lastRowNum = GetRowIndex(lastCellName);
            string firstColumn = GetColumnName(firstCellName);
            string lastColumn = GetColumnName(lastCellName);

            double sum = 0;

            // Iterate through the cells within the range and add their values to the sum.
            foreach (Row row in worksheet.Descendants<Row>().Where(r => r.RowIndex.Value >= firstRowNum && r.RowIndex.Value <= lastRowNum))
            {
                foreach (Cell cell in row)
                {
                    string columnName = GetColumnName(cell.CellReference.Value);
                    if (CompareColumn(columnName, firstColumn) >= 0 && CompareColumn(columnName, lastColumn) <= 0)
                    {
                        sum += double.Parse(cell.CellValue.Text);
                    }
                }
            }

            // Get the SharedStringTablePart and add the result to it.
            // If the SharedStringPart does not exist, create a new one.
            SharedStringTablePart shareStringPart;
            if (document.WorkbookPart.GetPartsOfType<SharedStringTablePart>().Count() > 0)
            {
                shareStringPart = document.WorkbookPart.GetPartsOfType<SharedStringTablePart>().First();
            }
            else
            {
                shareStringPart = document.WorkbookPart.AddNewPart<SharedStringTablePart>();
            }

            // Insert the result into the SharedStringTablePart.
            int index = InsertSharedStringItem("Result: " + sum, shareStringPart);

            Cell result = InsertCellInWorksheet(GetColumnName(resultCell), GetRowIndex(resultCell), worksheetPart);

            // Set the value of the cell.
            result.CellValue = new CellValue(index.ToString());
            result.DataType = new EnumValue<CellValues>(CellValues.SharedString);

            worksheetPart.Worksheet.Save();
        }
    }

    // Given a cell name, parses the specified cell to get the row index.
    private static uint GetRowIndex(string cellName)
    {
        // Create a regular expression to match the row index portion the cell name.
        Regex regex = new Regex(@"\d+");
        Match match = regex.Match(cellName);

        return uint.Parse(match.Value);
    }
    // Given a cell name, parses the specified cell to get the column name.
    private static string GetColumnName(string cellName)
    {
        // Create a regular expression to match the column name portion of the cell name.
        Regex regex = new Regex("[A-Za-z]+");
        Match match = regex.Match(cellName);

        return match.Value;
    }
    // Given two columns, compares the columns.
    private static int CompareColumn(string column1, string column2)
    {
        if (column1.Length > column2.Length)
        {
            return 1;
        }
        else if (column1.Length < column2.Length)
        {
            return -1;
        }
        else
        {
            return string.Compare(column1, column2, true);
        }
    }
    // Given text and a SharedStringTablePart, creates a SharedStringItem with the specified text 
    // and inserts it into the SharedStringTablePart. If the item already exists, returns its index.
    private static int InsertSharedStringItem(string text, SharedStringTablePart shareStringPart)
    {
        // If the part does not contain a SharedStringTable, create it.
        if (shareStringPart.SharedStringTable == null)
        {
            shareStringPart.SharedStringTable = new SharedStringTable();
        }

        int i = 0;
        foreach (SharedStringItem item in shareStringPart.SharedStringTable.Elements<SharedStringItem>())
        {
            if (item.InnerText == text)
            {
                // The text already exists in the part. Return its index.
                return i;
            }

            i++;
        }

        // The text does not exist in the part. Create the SharedStringItem.
        shareStringPart.SharedStringTable.AppendChild(new SharedStringItem(new DocumentFormat.OpenXml.Spreadsheet.Text(text)));
        shareStringPart.SharedStringTable.Save();

        return i;
    }
    // Given a column name, a row index, and a WorksheetPart, inserts a cell into the worksheet. 
    // If the cell already exists, returns it. 
    private static Cell InsertCellInWorksheet(string columnName, uint rowIndex, WorksheetPart worksheetPart)
    {
        Worksheet worksheet = worksheetPart.Worksheet;
        SheetData sheetData = worksheet.GetFirstChild<SheetData>();
        string cellReference = columnName + rowIndex;

        // If the worksheet does not contain a row with the specified row index, insert one.
        Row row;
        if (sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).Count() != 0)
        {
            row = sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).First();
        }
        else
        {
            row = new Row() { RowIndex = rowIndex };
            sheetData.Append(row);
        }

        // If there is not a cell with the specified column name, insert one.  
        if (row.Elements<Cell>().Where(c => c.CellReference.Value == columnName + rowIndex).Count() > 0)
        {
            return row.Elements<Cell>().Where(c => c.CellReference.Value == cellReference).First();
        }
        else
        {
            // Cells must be in sequential order according to CellReference. Determine where to insert the new cell.
            Cell refCell = null;
            foreach (Cell cell in row.Elements<Cell>())
            {
                if (string.Compare(cell.CellReference.Value, cellReference, true) > 0)
                {
                    refCell = cell;
                    break;
                }
            }

            Cell newCell = new Cell() { CellReference = cellReference };
            row.InsertBefore(newCell, refCell);

            worksheet.Save();
            return newCell;
        }
    }
```



## See also

- [Open XML SDK 2.5 class library reference](/office/open-xml/open-xml-sdk.md)
- [Language-Integrated Query (LINQ) (C#)](/dotnet/csharp/programming-guide/concepts/linq/)
- [Language-Integrated Query (LINQ) (Visual Basic)](/dotnet/visual-basic/programming-guide/concepts/linq/)
- [Lambda Expressions (C#)](/dotnet/csharp/language-reference/operators/lambda-expressions.md)
- [Lambda Expressions (Visual Basic)](/dotnet/visual-basic/programming-guide/language-features/procedures/lambda-expressions.md)

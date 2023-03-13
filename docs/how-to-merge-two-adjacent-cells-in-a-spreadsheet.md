# Merge two adjacent cells in a spreadsheet document (Open XML SDK)

This topic shows how to use the classes in the Open XML SDK 2.5 for
Office to merge two adjacent cells in a spreadsheet document
programmatically.

The following assembly directives are required to compile the code in
this topic.

```csharp
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using DocumentFormat.OpenXml;
    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.Spreadsheet;
    using System.Text.RegularExpressions;
```



--------------------------------------------------------------------------------
## Getting a SpreadsheetDocument Object 

In the Open XML SDK, the **[SpreadsheetDocument](https://msdn.microsoft.com/library/office/documentformat.openxml.packaging.spreadsheetdocument.aspx)** class represents an
Excel document package. To open and work with an Excel document, you
create an instance of the **SpreadsheetDocument** class from the document.
After you create the instance from the document, you can then obtain
access to the main workbook part that contains the worksheets. The text
in the document is represented in the package as XML using **SpreadsheetML** markup.

To create the class instance from the document that you call one of the
**[Open()](https://msdn.microsoft.com/library/office/documentformat.openxml.packaging.spreadsheetdocument.open.aspx)** overload methods. Several are
provided, each with a different signature. The sample code in this topic
uses the **[Open(String, Boolean)](https://msdn.microsoft.com/library/office/cc562356.aspx)** method with a
signature that requires two parameters. The first parameter takes a full
path string that represents the document that you want to open. The
second parameter is either **true** or **false** and represents whether you want the file to
be opened for editing. Any changes that you make to the document will
not be saved if this parameter is **false**.

The code that calls the **Open** method is
shown in the following **using** statement.

```csharp
    // Open the document for editing.
    using (SpreadsheetDocument document = SpreadsheetDocument.Open(docName, true)) 
    {
        // Insert other code here.
    }
```



The **using** statement provides a recommended
alternative to the typical .Open, .Save, .Close sequence. It ensures
that the **Dispose** method (internal method
used by the Open XML SDK to clean up resources) is automatically called
when the closing brace is reached. The block that follows the **using** statement establishes a scope for the
object that is created or named in the **using** statement, in this case **document**.


--------------------------------------------------------------------------------
## Basic Structure of a SpreadsheetML Document 

The basic document structure of a **SpreadsheetML** document consists of the **[Sheets](https://msdn.microsoft.com/library/office/documentformat.openxml.spreadsheet.sheets.aspx)** and **[Sheet](https://msdn.microsoft.com/library/office/documentformat.openxml.spreadsheet.sheet.aspx)** elements, which reference the
worksheets in the **Workbook**.
A separate XML file is created for each **Worksheet**.
For example, the **SpreadsheetML** for a
workbook that has two worksheets name MySheet1 and MySheet2 is located
in the Workbook.xml file and is shown in the following code example.

```xml
    <?xml version="1.0" encoding="UTF-8" standalone="yes" ?> 
    <workbook xmlns=https://schemas.openxmlformats.org/spreadsheetml/2006/main xmlns:r="https://schemas.openxmlformats.org/officeDocument/2006/relationships">
        <sheets>
            <sheet name="MySheet1" sheetId="1" r:id="rId1" /> 
            <sheet name="MySheet2" sheetId="2" r:id="rId2" /> 
        </sheets>
    </workbook>
```

The worksheet XML files contain one or more block level elements such as
**[SheetData](https://msdn.microsoft.com/library/office/documentformat.openxml.spreadsheet.sheetdata.aspx)**. **sheetData** represents the cell table and contains
one or more **[Row](https://msdn.microsoft.com/library/office/documentformat.openxml.spreadsheet.row.aspx)** elements. A **row** contains one or more **[Cell](https://msdn.microsoft.com/library/office/documentformat.openxml.spreadsheet.cell.aspx)** elements. Each cell contains a **[CellValue](https://msdn.microsoft.com/library/office/documentformat.openxml.spreadsheet.cellvalue.aspx)** element that represents the value
of the cell. For example, the SpreadsheetML for the first worksheet in a
workbook, that only has the value 100 in cell A1, is located in the
Sheet1.xml file and is shown in the following code example.

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

Using the Open XML SDK 2.5, you can create document structure and
content that uses strongly-typed classes that correspond to **SpreadsheetML** elements. You can find these
classes in the **DocumentFormat.OpenXML.Spreadsheet** namespace. The
following table lists the class names of the classes that correspond to
the **workbook**, **sheets**, **sheet**, **worksheet**, and **sheetData** elements.

| SpreadsheetML Element | Open XML SDK 2.5 Class | Description |
|---|---|---|
| workbook | DocumentFormat.OpenXML.Spreadsheet.Workbook | The root element for the main document part. |
| sheets | DocumentFormat. OpenXML.Spreadsheet.Sheets | The container for the block level structures such as sheet, fileVersion, and others specified in the [ISO/IEC 29500](https://www.iso.org/standard/71691.html) specification. |
| sheet | DocumentFormat.OpenXML.Spreadsheet.Sheet | A sheet that points to a sheet definition file. |
| worksheet | DocumentFormat.OpenXML.Spreadsheet.Worksheet | A sheet definition file that contains the sheet data. |
| sheetData | DocumentFormat.OpenXML.Spreadsheet.SheetData | The cell table, grouped together by rows. |
| row | DocumentFormat.OpenXml.Spreadsheet.Row | A row in the cell table. |
| c | DocumentFormat.OpenXml.Spreadsheet.Cell | A cell in a row. |
| v | DocumentFormat.OpenXml.Spreadsheet.CellValue | The value of a cell. |


--------------------------------------------------------------------------------
## How the Sample Code Works 

After you have opened the spreadsheet file for editing, the code
verifies that the specified cells exist, and if they do not exist, it
creates them by calling the **CreateSpreadsheetCellIfNotExist** method and append
it to the appropriate **[Row](https://msdn.microsoft.com/library/office/documentformat.openxml.spreadsheet.row.aspx)** object.

```csharp
    // Given a Worksheet and a cell name, verifies that the specified cell exists.
    // If it does not exist, creates a new cell. 
    private static void CreateSpreadsheetCellIfNotExist(Worksheet worksheet, string cellName)
    {
        string columnName = GetColumnName(cellName);
        uint rowIndex = GetRowIndex(cellName);

        IEnumerable<Row> rows = worksheet.Descendants<Row>().Where(r => r.RowIndex.Value == rowIndex);

        // If the Worksheet does not contain the specified row, create the specified row.
        // Create the specified cell in that row, and insert the row into the Worksheet.
        if (rows.Count() == 0)
        {
            Row row = new Row() { RowIndex = new UInt32Value(rowIndex) };
            Cell cell = new Cell() { CellReference = new StringValue(cellName) };
            row.Append(cell);
            worksheet.Descendants<SheetData>().First().Append(row);
            worksheet.Save();
        }
        else
        {
            Row row = rows.First();

            IEnumerable<Cell> cells = row.Elements<Cell>().Where(c => c.CellReference.Value == cellName);

            // If the row does not contain the specified cell, create the specified cell.
            if (cells.Count() == 0)
            {
                Cell cell = new Cell() { CellReference = new StringValue(cellName) };
                row.Append(cell);
                worksheet.Save();
            }
        }
    }
```



In order to get a column name, the code creates a new regular expression
to match the column name portion of the cell name. This regular
expression matches any combination of uppercase or lowercase letters.
For more information about regular expressions, see [Regular Expression
Language
Elements](https://msdn.microsoft.com/library/az24scfc.aspx). The
code gets the column name by calling the [Regex.Match](https://msdn2.microsoft.com/library/3zy662f6).

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



To get the row index, the code creates a new regular expression to match
the row index portion of the cell name. This regular expression matches
any combination of decimal digits. The following code creates a regular
expression to match the row index portion of the cell name, comprised of
decimal digits.

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



--------------------------------------------------------------------------------
## Sample Code 

The following code merges two adjacent cells in a **[SpreadsheetDocument](https://msdn.microsoft.com/library/office/documentformat.openxml.spreadsheet.row.aspx)** document package. When
merging two cells, only the content from one of the cells is preserved.
In left-to-right languages, the content in the upper-left cell is
preserved. In right-to-left languages, the content in the upper-right
cell is preserved. You can call the **MergeTwoCells** method in your program by using the
following code example, which merges the two cells B2 and C2 in a sheet
named "Jane," in a file named "Sheet9.xlsx."

```csharp
    string docName = @"C:\Users\Public\Documents\Sheet9.xlsx";
    string sheetName = "Jane";
    string cell1Name = "B2";
    string cell2Name = "C2";
    MergeTwoCells(docName, sheetName, cell1Name, cell2Name);
```



The following is the complete sample code in both C\# and Visual Basic.

```csharp
    // Given a document name, a worksheet name, and the names of two adjacent cells, merges the two cells.
    // When two cells are merged, only the content from one cell is preserved:
    // the upper-left cell for left-to-right languages or the upper-right cell for right-to-left languages.
    private static void MergeTwoCells(string docName, string sheetName, string cell1Name, string cell2Name)
    {
        // Open the document for editing.
        using (SpreadsheetDocument document = SpreadsheetDocument.Open(docName, true))
        {
            Worksheet worksheet = GetWorksheet(document, sheetName);
            if (worksheet == null || string.IsNullOrEmpty(cell1Name) || string.IsNullOrEmpty(cell2Name))
            {
                return;
            }

            // Verify if the specified cells exist, and if they do not exist, create them.
            CreateSpreadsheetCellIfNotExist(worksheet, cell1Name);
            CreateSpreadsheetCellIfNotExist(worksheet, cell2Name);

            MergeCells mergeCells;
            if (worksheet.Elements<MergeCells>().Count() > 0)
            {
                mergeCells = worksheet.Elements<MergeCells>().First();
            }
            else
            {
                mergeCells = new MergeCells();

                // Insert a MergeCells object into the specified position.
                if (worksheet.Elements<CustomSheetView>().Count() > 0)
                {
                    worksheet.InsertAfter(mergeCells, worksheet.Elements<CustomSheetView>().First());
                }
                else if (worksheet.Elements<DataConsolidate>().Count() > 0)
                {
                    worksheet.InsertAfter(mergeCells, worksheet.Elements<DataConsolidate>().First());
                }
                else if (worksheet.Elements<SortState>().Count() > 0)
                {
                    worksheet.InsertAfter(mergeCells, worksheet.Elements<SortState>().First());
                }
                else if (worksheet.Elements<AutoFilter>().Count() > 0)
                {
                    worksheet.InsertAfter(mergeCells, worksheet.Elements<AutoFilter>().First());
                }
                else if (worksheet.Elements<Scenarios>().Count() > 0)
                {
                    worksheet.InsertAfter(mergeCells, worksheet.Elements<Scenarios>().First());
                }
                else if (worksheet.Elements<ProtectedRanges>().Count() > 0)
                {
                    worksheet.InsertAfter(mergeCells, worksheet.Elements<ProtectedRanges>().First());
                }
                else if (worksheet.Elements<SheetProtection>().Count() > 0)
                {
                    worksheet.InsertAfter(mergeCells, worksheet.Elements<SheetProtection>().First());
                }
                else if (worksheet.Elements<SheetCalculationProperties>().Count() > 0)
                {
                    worksheet.InsertAfter(mergeCells, worksheet.Elements<SheetCalculationProperties>().First());
                }
                else
                {
                    worksheet.InsertAfter(mergeCells, worksheet.Elements<SheetData>().First());
                }
            }

            // Create the merged cell and append it to the MergeCells collection.
            MergeCell mergeCell = new MergeCell() { Reference = new StringValue(cell1Name + ":" + cell2Name) };
            mergeCells.Append(mergeCell);

            worksheet.Save();
        }
    }
    // Given a Worksheet and a cell name, verifies that the specified cell exists.
    // If it does not exist, creates a new cell. 
    private static void CreateSpreadsheetCellIfNotExist(Worksheet worksheet, string cellName)
    {
        string columnName = GetColumnName(cellName);
        uint rowIndex = GetRowIndex(cellName);

        IEnumerable<Row> rows = worksheet.Descendants<Row>().Where(r => r.RowIndex.Value == rowIndex);

        // If the Worksheet does not contain the specified row, create the specified row.
        // Create the specified cell in that row, and insert the row into the Worksheet.
        if (rows.Count() == 0)
        {
            Row row = new Row() { RowIndex = new UInt32Value(rowIndex) };
            Cell cell = new Cell() { CellReference = new StringValue(cellName) };
            row.Append(cell);
            worksheet.Descendants<SheetData>().First().Append(row);
            worksheet.Save();
        }
        else
        {
            Row row = rows.First();

            IEnumerable<Cell> cells = row.Elements<Cell>().Where(c => c.CellReference.Value == cellName);

            // If the row does not contain the specified cell, create the specified cell.
            if (cells.Count() == 0)
            {
                Cell cell = new Cell() { CellReference = new StringValue(cellName) };
                row.Append(cell);
                worksheet.Save();
            }
        }
    }

    // Given a SpreadsheetDocument and a worksheet name, get the specified worksheet.
    private static Worksheet GetWorksheet(SpreadsheetDocument document, string worksheetName)
    {
        IEnumerable<Sheet> sheets = document.WorkbookPart.Workbook.Descendants<Sheet>().Where(s => s.Name == worksheetName);
        WorksheetPart worksheetPart = (WorksheetPart)document.WorkbookPart.GetPartById(sheets.First().Id);
        if (sheets.Count() == 0)
            return null;
        else
            return worksheetPart.Worksheet;
    }

    // Given a cell name, parses the specified cell to get the column name.
    private static string GetColumnName(string cellName)
    {
        // Create a regular expression to match the column name portion of the cell name.
        Regex regex = new Regex("[A-Za-z]+");
        Match match = regex.Match(cellName);

        return match.Value;
    }
    // Given a cell name, parses the specified cell to get the row index.
    private static uint GetRowIndex(string cellName)
    {
        // Create a regular expression to match the row index portion the cell name.
        Regex regex = new Regex(@"\d+");
        Match match = regex.Match(cellName);

        return uint.Parse(match.Value);
    }
```



--------------------------------------------------------------------------------
## See also 



- [Open XML SDK 2.5 class library reference](/office/open-xml/open-xml-sdk)

[Language-Integrated Query (LINQ)](https://msdn.microsoft.com/library/bb397926.aspx)

[Lambda Expressions](https://msdn.microsoft.com/library/bb531253.aspx)

[Lambda Expressions (C\# Programming Guide)](https://msdn.microsoft.com/library/bb397687.aspx)

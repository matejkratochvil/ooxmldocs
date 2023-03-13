# Retrieve a list of the worksheets in a spreadsheet document (Open XML SDK)

This topic shows how to use the classes in the Open XML SDK 2.5 for
Office to programmatically retrieve a list of the worksheets in a
Microsoft Excel 2010 or Microsoft Excel 2013 workbook, without loading
the document into Excel. It contains an example **GetAllWorksheets** method to illustrate this task.

To use the sample code in this topic, you must install the [Open XML SDK 2.5](https://www.nuget.org/packages/DocumentFormat.OpenXml/2.5.0). You
must explicitly reference the following assemblies in your project:

- WindowsBase

- DocumentFormat.OpenXml (installed by the Open XML SDK)

You must also use the following **using**
directives or **Imports** statements to compile
the code in this topic.

```csharp
    using System;
    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.Spreadsheet;
```



--------------------------------------------------------------------------------

## GetAllWorksheets Method

You can use the **GetAllWorksheets** method,
which is shown in the following code, to retrieve a list of the
worksheets in a workbook. The **GetAllWorksheets** method accepts a single
parameter, a string that indicates the path of the file that you want to
examine.

```csharp
    public static Sheets GetAllWorksheets(string fileName)
```



The method works with the workbook you specify, returning an instance of
the **[Sheets](https://msdn.microsoft.com/library/office/documentformat.openxml.spreadsheet.sheets.aspx)** object, from which you can retrieve
a reference to each **[Sheet](https://msdn.microsoft.com/library/office/documentformat.openxml.spreadsheet.sheet.aspx)** object.

--------------------------------------------------------------------------------

## Calling the GetAllWorksheets Method

To call the **GetAllWorksheets** method, pass
the required value, as shown in the following code.

```csharp
    const string DEMOFILE = @"C:\Users\Public\Documents\SampleWorkbook.xlsx";

    static void Main(string[] args)
    {
        var results = GetAllWorksheets(DEMOFILE);
        foreach (Sheet item in results)
        {
            Console.WriteLine(item.Name);
        }
    }
```



--------------------------------------------------------------------------------

## How the Code Works

The sample method, **GetAllWorksheets**,
creates a variable that will contain a reference to the **Sheets** collection of the workbook. At the end of
its work, the method returns the variable, which contains either a
reference to the **Sheets** collection, or
null/Nothing if there were no sheets (this cannot occur in a well-formed
workbook).

```csharp
    Sheets theSheets = null;
    // Code removed here…
    return theSheets;
```



The code then continues by opening the document in read-only mode, and
retrieving a reference to the **[WorkbookPart](https://msdn.microsoft.com/library/office/documentformat.openxml.packaging.spreadsheetdocument.workbookpart.aspx)**.

```csharp
    using (SpreadsheetDocument document = 
        SpreadsheetDocument.Open(fileName, false))
    {
        WorkbookPart wbPart = document.WorkbookPart;
        // Code removed here.
    }
```



To get access to the **[Workbook](https://msdn.microsoft.com/library/office/documentformat.openxml.spreadsheet.workbook.aspx)** object, the code retrieves the value of the **[Workbook](https://msdn.microsoft.com/library/office/documentformat.openxml.packaging.workbookpart.workbook.aspx)** property from the **WorkbookPart**, and then retrieves a reference to the **Sheets** object from the **[Sheets](https://msdn.microsoft.com/library/office/documentformat.openxml.spreadsheet.workbook.sheets.aspx)** property of the **Workbook**. The **Sheets** object contains the collection of **[Sheet](https://msdn.microsoft.com/library/office/documentformat.openxml.spreadsheet.sheet.aspx)** objects that provide the method's return value.

```csharp
    theSheets = wbPart.Workbook.Sheets;
```



--------------------------------------------------------------------------------

## Sample Code

The following is the complete **GetAllWorksheets** code sample in C\# and Visual
Basic.

```csharp
    using System;
    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.Spreadsheet;

    namespace GetAllWorkheets
    {
        class Program
        {
            const string DEMOFILE = 
                @"C:\Users\Public\Documents\SampleWorkbook.xlsx";

            static void Main(string[] args)
            {
                var results = GetAllWorksheets(DEMOFILE);
                foreach (Sheet item in results)
                {
                    Console.WriteLine(item.Name);
                }
            }

            // Retrieve a List of all the sheets in a workbook.
            // The Sheets class contains a collection of 
            // OpenXmlElement objects, each representing one of 
            // the sheets.
            public static Sheets GetAllWorksheets(string fileName)
            {
                Sheets theSheets = null;

                using (SpreadsheetDocument document = 
                    SpreadsheetDocument.Open(fileName, false))
                {
                    WorkbookPart wbPart = document.WorkbookPart;
                    theSheets = wbPart.Workbook.Sheets;
                }
                return theSheets;
            }
        }
    }
```



--------------------------------------------------------------------------------

## See also

- [Open XML SDK 2.5 class library reference](/office/open-xml/open-xml-sdk)



# Retrieve a list of the hidden worksheets in a spreadsheet document (Open XML SDK)

This topic shows how to use the classes in the Open XML SDK 2.5 for Office to programmatically retrieve a list of hidden worksheets in a Microsoft Excel 2010 or Microsoft Excel 2010 workbook, without loading the document into Excel. It contains an example **GetHiddenSheets** method to illustrate this task.

To use the sample code in this topic, you must install the [Open XML SDK 2.5](https://www.nuget.org/packages/DocumentFormat.OpenXml/2.5.0). You must explicitly reference the following assemblies in your project:

- WindowsBase
- DocumentFormat.OpenXml (installed by the Open XML SDK)

You must also use the following **using** directives or **Imports** statements to compile the code in this topic.

```csharp
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.Spreadsheet;
```



## GetHiddenSheets method

You can use the **GetHiddenSheets** method, which is shown in the following code, to retrieve a list of the hidden worksheets in a workbook. The **GetHiddenSheets** method accepts a single parameter, a string that indicates the path of the file that you want to examine.

```csharp
    public static List<Sheet> GetHiddenSheets(string fileName)
```



The method works with the workbook you specify, filling a **[List\<T\>](https://msdn2.microsoft.com/library/6sh2ey19)** instance with a reference to each hidden **Sheet** object.

## Calling the GetHiddenSheets method

The method returns a generic list that contains information about the individual hidden **Sheet** objects. To call the **GetHiddenWorksheets** method, pass the required parameter value, as shown in the following code.

```csharp
    // Revise this path to the location of a file that contains hidden worksheets.
    const string DEMOPATH = 
        @"C:\Users\Public\Documents\HiddenSheets.xlsx";
    List<Sheet> sheets = GetHiddenSheets(DEMOPATH);
    foreach (var sheet in sheets)
    {
        Console.WriteLine(sheet.Name);
    }
```



## How the code works

The following code starts by creating a generic list that will contain information about the hidden worksheets.

```csharp
    List<Sheet> returnVal = new List<Sheet>();
```



Next, the following code opens the specified workbook by using the **SpreadsheetDocument.Open** method and indicating that the document should be open for read-only access (the final **false** parameter value). Given the open workbook, the code uses the **WorkbookPart** property to navigate to the main workbook part, storing the reference in a variable named **wbPart**.

```csharp
    using (SpreadsheetDocument document = 
        SpreadsheetDocument.Open(fileName, false))
    {
        WorkbookPart wbPart = document.WorkbookPart;
        // Code removed here… 
    }
    return returnVal;
```



## Retrieve the collection of worksheets

The **WorkbookPart** class provides a **Workbook** property, which in turn contains the XML content of the workbook. Although the Open XML SDK 2.5 provides the **Sheets** property, which returns a collection of the **Sheet** parts, all the information that you need is provided by the **Sheet** elements within the **Workbook** XML content.
The following code uses the **Descendants** generic method of the **Workbook** object to retrieve a collection of **Sheet** objects that contain information about all the sheet child elements of the workbook's XML content.

```csharp
    var sheets = wbPart.Workbook.Descendants<Sheet>();
```



## Retrieve hidden sheets

It's important to be aware that Excel supports two levels of worksheets. You can hide a worksheet by using the Excel user interface by right-clicking the worksheets tab and opting to hide the worksheet.
For these worksheets, the **State** property of the **Sheet** object contains an enumerated value of **Hidden**. You can also make a worksheet very hidden by writing code (either in VBA or in another language) that sets the sheet's **Visible** property to the enumerated value **xlSheetVeryHidden**. For worksheets hidden in this manner, the **State** property of the **Sheet** object contains the enumerated value **VeryHidden**.

Given the collection that contains information about all the sheets, the following code uses the **[Where](https://msdn2.microsoft.com/library/bb301979)** function to filter the collection so that it contains only the sheets in which the **State** property is not null. If the **State** property is not null, the code looks for the **Sheet** objects in which the **State** property as a value, and where the value is either **SheetStateValues.Hidden** or **SheetStateValues.VeryHidden**.

```csharp
    var hiddenSheets = sheets.Where((item) => item.State != null && 
        item.State.HasValue && 
        (item.State.Value == SheetStateValues.Hidden || 
        item.State.Value == SheetStateValues.VeryHidden));
```



Finally, the following code calls the **[ToList\<TSource\>](https://msdn2.microsoft.com/library/bb342261)** method to execute the LINQ query that retrieves the list of hidden sheets, placing the result into the return value for the function.

```csharp
    returnVal = hiddenSheets.ToList();
```



## Sample code

The following is the complete **GetHiddenSheets** code sample in C\# and Visual Basic.

```csharp
    public static List<Sheet> GetHiddenSheets(string fileName)
    {
        List<Sheet> returnVal = new List<Sheet>();

        using (SpreadsheetDocument document = 
            SpreadsheetDocument.Open(fileName, false))
        {
            WorkbookPart wbPart = document.WorkbookPart;
            var sheets = wbPart.Workbook.Descendants<Sheet>();

            // Look for sheets where there is a State attribute defined, 
            // where the State has a value,
            // and where the value is either Hidden or VeryHidden.
            var hiddenSheets = sheets.Where((item) => item.State != null &&
                item.State.HasValue &&
                (item.State.Value == SheetStateValues.Hidden ||
                item.State.Value == SheetStateValues.VeryHidden));

            returnVal = hiddenSheets.ToList();
        }
        return returnVal;
    }
```



## See also

- [Open XML SDK 2.5 class library reference](/office/open-xml/open-xml-sdk)



# Add tables to word processing documents (Open XML SDK)

This topic shows how to use the classes in the Open XML SDK 2.5 for Office to programmatically add a table to a word processing document. It contains an example **AddTable** method to illustrate this task.

To use the sample code in this topic, you must install the [Open XML SDK 2.5](https://www.nuget.org/packages/DocumentFormat.OpenXml/2.5.0). You must explicitly reference the following assemblies in your project:

- WindowsBase
- DocumentFormat.OpenXml (installed by the Open XML SDK)

You must also use the following **using** directives or **Imports** statements to compile the code in this topic.

```csharp
    using DocumentFormat.OpenXml;
    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.Wordprocessing;
```



## AddTable method

You can use the **AddTable** method to add a simple table to a word processing document. The **AddTable** method accepts two parameters, indicating the following:

- The name of the document to modify (string).

- A two-dimensional array of strings to insert into the document as a
    table.

```csharp
    public static void AddTable(string fileName, string[,] data)
```



## Call the AddTable method

The **AddTable** method modifies the document you specify, adding a table that contains the information in the two-dimensional array that you provide. To call the method, pass both of the parameter values, as shown in the following code.

```csharp
    const string fileName = @"C:\Users\Public\Documents\AddTable.docx";
    AddTable(fileName, new string[,] 
        { { "Texas", "TX" }, 
        { "California", "CA" }, 
        { "New York", "NY" }, 
        { "Massachusetts", "MA" } }
        );
```



## How the code works

The following code starts by opening the document, using the [WordprocessingDocument.Open](/dotnet/api/documentformat.openxml.packaging.wordprocessingdocument.open) method and indicating that the document should be open for read/write access (the final **true** parameter value). Next the code retrieves a reference to the root element of the main document part, using the [Document](/dotnet/api/documentformat.openxml.packaging.maindocumentpart.document) property of the [MainDocumentPart](/dotnet/api/documentformat.openxml.packaging.wordprocessingdocument.maindocumentpart) of the word processing document.

```csharp
    using (var document = WordprocessingDocument.Open(fileName, true))
    {
        var doc = document.MainDocumentPart.Document;
        // Code removed here…
    }
```



## Create the table object and set its properties

Before you can insert a table into a document, you must create the [Table](/dotnet/api/documentformat.openxml.wordprocessing.table) object and set its properties. To set a table's properties, you create and supply values for a [TableProperties](/dotnet/api/documentformat.openxml.wordprocessing.tableproperties) object. The **TableProperties** class provides many table-oriented properties, like [Shading](/dotnet/api/documentformat.openxml.wordprocessing.tableproperties.shading), [TableBorders](/dotnet/api/documentformat.openxml.wordprocessing.tableproperties.tableborders), [TableCaption](/dotnet/api/documentformat.openxml.wordprocessing.tableproperties.tablecaption), [TableCellSpacing](/dotnet/api/documentformat.openxml.wordprocessing.tableproperties.tablecellspacing), [TableJustification](/dotnet/api/documentformat.openxml.wordprocessing.tableproperties.tablejustification), and more. The sample method includes the following code.

```csharp
    Table table = new Table();

    TableProperties props = new TableProperties(
        new TableBorders(
        new TopBorder
        {
            Val = new EnumValue<BorderValues>(BorderValues.Single),
            Size = 12
        },
        new BottomBorder
        {
          Val = new EnumValue<BorderValues>(BorderValues.Single),
          Size = 12
        },
        new LeftBorder
        {
          Val = new EnumValue<BorderValues>(BorderValues.Single),
          Size = 12
        },
        new RightBorder
        {
          Val = new EnumValue<BorderValues>(BorderValues.Single),
          Size = 12
        },
        new InsideHorizontalBorder
        {
          Val = new EnumValue<BorderValues>(BorderValues.Single),
          Size = 12
        },
        new InsideVerticalBorder
        {
          Val = new EnumValue<BorderValues>(BorderValues.Single),
          Size = 12
    }));

    table.AppendChild<TableProperties>(props);
```



The constructor for the **TableProperties** class allows you to specify as many child elements as you like (much like the [XElement](/dotnet/api/system.xml.linq.xelement) constructor). In this case, the code creates [TopBorder](/dotnet/api/documentformat.openxml.wordprocessing.topborder), [BottomBorder](/dotnet/api/documentformat.openxml.wordprocessing.bottomborder), [LeftBorder](/dotnet/api/documentformat.openxml.wordprocessing.leftborder), [RightBorder](/dotnet/api/documentformat.openxml.wordprocessing.rightborder), [InsideHorizontalBorder](/dotnet/api/documentformat.openxml.wordprocessing.insidehorizontalborder), and [InsideVerticalBorder](/dotnet/api/documentformat.openxml.wordprocessing.insideverticalborder) child elements, each describing one of the border elements for the table. For each element, the code sets the **Val** and **Size** properties as part of calling the constructor. Setting the size is simple, but setting the **Val** property requires a bit more effort: this property, for this particular object, represents the border style, and you must set it to an enumerated value. To do that, create an instance of the [EnumValue\<T\>](/dotnet/api/documentformat.openxml.enumvalue-1) generic type, passing the specific border type ([Single](/dotnet/api/documentformat.openxml.wordprocessing.bordervalues) as a parameter to the constructor. Once the code has set all the table border value it needs to set, it calls the [AppendChild\<T\>](/dotnet/api/documentformat.openxml.openxmlelement.appendchild) method of the table, indicating that the generic type is [TableProperties](/dotnet/api/ ocumentformat.openxml.wordprocessing.tableproperties)—that is, it is appending an instance of the **TableProperties** class, using the variable **props** as the value.

## Fill the table with data

Given that table and its properties, now it is time to fill the table with data. The sample procedure iterates first through all the rows of data in the array of strings that you specified, creating a new [TableRow](/dotnet/api/documentformat.openxml.wordprocessing.tablerow) instance for each row of data. The following code leaves out the details of filling in the row with data, but it shows how you create and append the row to the table:

```csharp
    for (var i = 0; i <= data.GetUpperBound(0); i++)
    {
        var tr = new TableRow();
        // Code removed here…
        table.Append(tr);
    }
```



For each row, the code iterates through all the columns in the array of strings you specified. For each column, the code creates a new [TableCell](/dotnet/api/documentformat.openxml.wordprocessing.tablecell) object, fills it with data, and appends it to the row. The following code leaves out the details of filling each cell with data, but it shows how you create and append the column to the table:

```csharp
    for (var j = 0; j <= data.GetUpperBound(1); j++)
    {
        var tc = new TableCell();
        // Code removed here…
        tr.Append(tc);
    }
```



Next, the code does the following:

- Creates a new [Text](/dotnet/api/documentformat.openxml.wordprocessing.text) object that contains a value from the array of strings.
- Passes the [Text](/dotnet/api/documentformat.openxml.wordprocessing.text) object to the constructor for a new [Run](/dotnet/api/documentformat.openxml.wordprocessing.run) object.
- Passes the [Run](/dotnet/api/documentformat.openxml.wordprocessing.run) object to the constructor for a new [Paragraph](/dotnet/api/documentformat.openxml.wordprocessing.paragraph) object.
- Passes the [Paragraph](/dotnet/api/documentformat.openxml.wordprocessing.paragraph) object to the [Append](/dotnet/api/documentformat.openxml.openxmlelement.append)method of the cell.

In other words, the following code appends the text to the new **TableCell** object.

```csharp
    tc.Append(new Paragraph(new Run(new Text(data[i, j]))));
```



The code then appends a new [TableCellProperties](/dotnet/api/documentformat.openxml.wordprocessing.tablecellproperties) object to the cell. This **TableCellProperties** object, like the **TableProperties** object you already saw, can accept as many objects in its constructor as you care to supply. In this case, the code passes only a new [TableCellWidth](/dotnet/api/documentformat.openxml.wordprocessing.tablecellwidth) object, with its [Type](/dotnet/api/documentformat.openxml.wordprocessing.tablewidthtype.type) property set to [Auto](/dotnet/api/documentformat.openxml.wordprocessing.tablewidthunitvalues) (so that the table automatically sizes the width of each column).

```csharp
    // Assume you want columns that are automatically sized.
    tc.Append(new TableCellProperties(
        new TableCellWidth { Type = TableWidthUnitValues.Auto }));
```



## Finish up

The following code concludes by appending the table to the body of the document, and then saving the document.

```csharp
    doc.Body.Append(table);
    doc.Save();
```



## Sample Code

The following is the complete **AddTable** code sample in C\# and Visual Basic.

```csharp
    // Take the data from a two-dimensional array and build a table at the 
    // end of the supplied document.
    public static void AddTable(string fileName, string[,] data)
    {
        using (var document = WordprocessingDocument.Open(fileName, true))
        {

            var doc = document.MainDocumentPart.Document;

            Table table = new Table();

            TableProperties props = new TableProperties(
                new TableBorders(
                new TopBorder
                {
                    Val = new EnumValue<BorderValues>(BorderValues.Single),
                    Size = 12
                },
                new BottomBorder
                {
                  Val = new EnumValue<BorderValues>(BorderValues.Single),
                  Size = 12
                },
                new LeftBorder
                {
                  Val = new EnumValue<BorderValues>(BorderValues.Single),
                  Size = 12
                },
                new RightBorder
                {
                  Val = new EnumValue<BorderValues>(BorderValues.Single),
                  Size = 12
                },
                new InsideHorizontalBorder
                {
                  Val = new EnumValue<BorderValues>(BorderValues.Single),
                  Size = 12
                },
                new InsideVerticalBorder
                {
                  Val = new EnumValue<BorderValues>(BorderValues.Single),
                  Size = 12
            }));

            table.AppendChild<TableProperties>(props);

            for (var i = 0; i <= data.GetUpperBound(0); i++)
            {
                var tr = new TableRow();
                for (var j = 0; j <= data.GetUpperBound(1); j++)
                {
                    var tc = new TableCell();
                    tc.Append(new Paragraph(new Run(new Text(data[i, j]))));
                    
                    // Assume you want columns that are automatically sized.
                    tc.Append(new TableCellProperties(
                        new TableCellWidth { Type = TableWidthUnitValues.Auto }));
                    
                    tr.Append(tc);
                }
                table.Append(tr);
            }
            doc.Body.Append(table);
            doc.Save();
        }
    }
```



## See also

- [Open XML SDK 2.5 class library reference](/office/open-xml/open-xml-sdk)


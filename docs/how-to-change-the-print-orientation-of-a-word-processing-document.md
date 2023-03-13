

# Change the print orientation of a word processing document

This topic shows how to use the classes in the Open XML SDK 2.5 for
Office to programmatically set the print orientation of a Microsoft Word
2010 or Microsoft Word 2013 document. It contains an example
**SetPrintOrientation** method to illustrate this task.

To use the sample code in this topic, you must install the [Open XML SDK
2.5](https://www.nuget.org/packages/DocumentFormat.OpenXml/2.5.0). You
must explicitly reference the following assemblies in your project:

- WindowsBase

- DocumentFormat.OpenXml (installed by the Open XML SDK)

You must also use the following **using**
directives or **Imports** statements to compile
the code in this topic.

```csharp
    using System.Linq;
    using DocumentFormat.OpenXml;
    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.Wordprocessing;
```



-----------------------------------------------------------------------------

## SetPrintOrientation Method

You can use the **SetPrintOrientation** method
to change the print orientation of a word processing document. The
method accepts two parameters that indicate the name of the document to
modify (string) and the new print orientation ([PageOrientationValues](https://msdn.microsoft.com/library/office/documentformat.openxml.wordprocessing.pageorientationvalues.aspx)).

The following code shows the **SetPrintOrientation** method.

```csharp
    public static void SetPrintOrientation(
      string fileName, PageOrientationValues newOrientation)
```



For each section in the document, if the new orientation differs from
the section's current print orientation, the code modifies the print
orientation for the section. In addition, the code must manually update
the width, height, and margins for each section.

-----------------------------------------------------------------------------

## Calling the Sample SetPrintOrientation Method

To call the sample **SetPrintOrientation**
method, pass a string that contains the name of the file to convert. The
following code shows an example method call.

```csharp
    SetPrintOrientation(@"C:\Users\Public\Documents\ChangePrintOrientation.docx", 
        PageOrientationValues.Landscape);
```



-----------------------------------------------------------------------------

## How the Code Works

The following code first opens the document by using the [Open](https://msdn.microsoft.com/library/office/documentformat.openxml.packaging.wordprocessingdocument.open.aspx) method and sets the **isEditable** parameter to
**true** to indicate that the document should
be read/write. The code maintains a Boolean variable that tracks whether
the document has changed (so that it can save the document later, if the
document has changed). The code retrieves a reference to the main
document part, and then uses that reference to retrieve a collection of
all of the descendants of type [SectionProperties](https://msdn.microsoft.com/library/office/documentformat.openxml.wordprocessing.sectionproperties.aspx) within the content of the
document. Later code will use this collection to set the orientation for
each section in turn.

```csharp
    using (var document = 
        WordprocessingDocument.Open(fileName, true))
    {
        bool documentChanged = false;

        var docPart = document.MainDocumentPart;
        var sections = docPart.Document.Descendants<SectionProperties>();
        // Code removed here...
    }
```



-----------------------------------------------------------------------------

## Iterating Through All the Sections

The next block of code iterates through all the sections in the collection of **SectionProperties** elements. For each section, the code initializes a variable that tracks whether the page orientation for the section was changed so the code can update the page size and margins. (If the new orientation matches the original orientation, the code will not update the page.) The code continues by retrieving a reference to the first [PageSize](https://msdn.microsoft.com/library/office/documentformat.openxml.wordprocessing.pagesize.aspx) descendant of the **SectionProperties** element. If the reference is not null, the code updates the orientation as required.

```csharp
    foreach (SectionProperties sectPr in sections)
    {
        bool pageOrientationChanged = false;

        PageSize pgSz = sectPr.Descendants<PageSize>().FirstOrDefault();
        if (pgSz != null)
        {
            // Code removed here...
        }
    }
```



-----------------------------------------------------------------------------

## Setting the Orientation for the Section

The next block of code first checks whether the [Orient](https://msdn.microsoft.com/library/office/documentformat.openxml.wordprocessing.pagesize.orient.aspx) property of the **PageSize** element exists. As with many properties
of Open XML elements, the property or attribute might not exist yet. In
that case, retrieving the property returns a null reference. By default,
if the property does not exist, and the new orientation is Portrait, the
code will not update the page. If the **Orient** property already exists, and its value
differs from the new orientation value supplied as a parameter to the
method, the code sets the **Value** property of
the **Orient** property, and sets both the
**pageOrientationChanged** and the **documentChanged** flags. (The code uses the **pageOrientationChanged** flag to determine whether it
must update the page size and margins. It uses the **documentChanged** flag to determine whether it must
save the document at the end.)

> [!NOTE]
> If the code must create the **Orient** property, it must also create the value to store in the property, as a new [EnumValue\<T\>](https://msdn.microsoft.com/library/office/cc801792.aspx) instance, supplying the new orientation in the **EnumValue** constructor.

```csharp
    if (pgSz.Orient == null)
    {
        if (newOrientation != PageOrientationValues.Portrait)
        {
            pageOrientationChanged = true;
            documentChanged = true;
            pgSz.Orient = 
                new EnumValue<PageOrientationValues>(newOrientation);
        }
    }
    else
    {
        if (pgSz.Orient.Value != newOrientation)
        {
            pgSz.Orient.Value = newOrientation;
            pageOrientationChanged = true;
            documentChanged = true;
        }
    }
```



-----------------------------------------------------------------------------

## Updating the Page Size

At this point in the code, the page orientation may have changed. If so,
the code must complete two more tasks. It must update the page size, and
update the page margins for the section. The first task is easy—the
following code just swaps the page height and width, storing the values
in the **PageSize** element.

```csharp
    if (pageOrientationChanged)
    {
        // Changing the orientation is not enough. You must also 
        // change the page size.
        var width = pgSz.Width;
        var height = pgSz.Height;
        pgSz.Width = height;
        pgSz.Height = width;
        // Code removed here...
    }
```



-----------------------------------------------------------------------------

## Updating the Margins

The next step in the sample procedure handles margins for the section.
If the page orientation has changed, the code must rotate the margins to
match. To do so, the code retrieves a reference to the [PageMargin](https://msdn.microsoft.com/library/office/documentformat.openxml.wordprocessing.pagemargin.aspx) element for the section. If the
element exists, the code rotates the margins. Note that the code rotates
the margins by 90 degrees—some printers rotate the margins by 270
degrees instead and you could modify the code to take that into account.
Also be aware that the [Top](https://msdn.microsoft.com/library/office/documentformat.openxml.wordprocessing.pagemargin.top.aspx) and [Bottom](https://msdn.microsoft.com/library/office/documentformat.openxml.wordprocessing.pagemargin.bottom.aspx) properties of the **PageMargin** object are signed values, and the
[Left](https://msdn.microsoft.com/library/office/documentformat.openxml.wordprocessing.pagemargin.left.aspx) and [Right](https://msdn.microsoft.com/library/office/documentformat.openxml.wordprocessing.pagemargin.right.aspx) properties are unsigned values. The
code must convert between the two types of values as it rotates the
margin settings, as shown in the following code.

```csharp
    PageMargin pgMar = 
        sectPr.Descendants<PageMargin>().FirstOrDefault();
    if (pgMar != null)
    {
        var top = pgMar.Top.Value;
        var bottom = pgMar.Bottom.Value;
        var left = pgMar.Left.Value;
        var right = pgMar.Right.Value;

        pgMar.Top = new Int32Value((int)left);
        pgMar.Bottom = new Int32Value((int)right);
        pgMar.Left = 
            new UInt32Value((uint)System.Math.Max(0, bottom));
        pgMar.Right = 
            new UInt32Value((uint)System.Math.Max(0, top));
    }
```



-----------------------------------------------------------------------------

## Saving the Document

After all the modifications, the code determines whether the document
has changed. If the document has changed, the code saves it.

```csharp
    if (documentChanged)
    {
        docPart.Document.Save();
    }
```



-----------------------------------------------------------------------------

## Sample Code

The following is the complete **SetPrintOrientation** code sample in C\# and Visual
Basic.

```csharp
    using System.Linq;
    using DocumentFormat.OpenXml;
    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.Wordprocessing;

    namespace ChangePrintOrientation
    {
        class Program
        {
            static void Main(string[] args)
            {
                SetPrintOrientation(@"C:\Users\Public\Documents\ChangePrintOrientation.docx", 
                    PageOrientationValues.Landscape);
            }

            // Given a document name, set the print orientation for 
            // all the sections of the document.
            public static void SetPrintOrientation(
              string fileName, PageOrientationValues newOrientation)
            {
                using (var document = 
                    WordprocessingDocument.Open(fileName, true))
                {
                    bool documentChanged = false;

                    var docPart = document.MainDocumentPart;
                    var sections = docPart.Document.Descendants<SectionProperties>();

                    foreach (SectionProperties sectPr in sections)
                    {
                        bool pageOrientationChanged = false;

                        PageSize pgSz = sectPr.Descendants<PageSize>().FirstOrDefault();
                        if (pgSz != null)
                        {
                            // No Orient property? Create it now. Otherwise, just 
                            // set its value. Assume that the default orientation 
                            // is Portrait.
                            if (pgSz.Orient == null)
                            {
                                // Need to create the attribute. You do not need to 
                                // create the Orient property if the property does not 
                                // already exist, and you are setting it to Portrait. 
                                // That is the default value.
                                if (newOrientation != PageOrientationValues.Portrait)
                                {
                                    pageOrientationChanged = true;
                                    documentChanged = true;
                                    pgSz.Orient = 
                                        new EnumValue<PageOrientationValues>(newOrientation);
                                }
                            }
                            else
                            {
                                // The Orient property exists, but its value
                                // is different than the new value.
                                if (pgSz.Orient.Value != newOrientation)
                                {
                                    pgSz.Orient.Value = newOrientation;
                                    pageOrientationChanged = true;
                                    documentChanged = true;
                                }
                            }

                            if (pageOrientationChanged)
                            {
                                // Changing the orientation is not enough. You must also 
                                // change the page size.
                                var width = pgSz.Width;
                                var height = pgSz.Height;
                                pgSz.Width = height;
                                pgSz.Height = width;

                                PageMargin pgMar = 
                                    sectPr.Descendants<PageMargin>().FirstOrDefault();
                                if (pgMar != null)
                                {
                                    // Rotate margins. Printer settings control how far you 
                                    // rotate when switching to landscape mode. Not having those
                                    // settings, this code rotates 90 degrees. You could easily
                                    // modify this behavior, or make it a parameter for the 
                                    // procedure.
                                    var top = pgMar.Top.Value;
                                    var bottom = pgMar.Bottom.Value;
                                    var left = pgMar.Left.Value;
                                    var right = pgMar.Right.Value;

                                    pgMar.Top = new Int32Value((int)left);
                                    pgMar.Bottom = new Int32Value((int)right);
                                    pgMar.Left = 
                                        new UInt32Value((uint)System.Math.Max(0, bottom));
                                    pgMar.Right = 
                                        new UInt32Value((uint)System.Math.Max(0, top));
                                }
                            }
                        }
                    }
                    if (documentChanged)
                    {
                        docPart.Document.Save();
                    }
                }
            }
        }
    }
```



-----------------------------------------------------------------------------

## See also

- [Open XML SDK 2.5 class library reference](/office/open-xml/open-xml-sdk)

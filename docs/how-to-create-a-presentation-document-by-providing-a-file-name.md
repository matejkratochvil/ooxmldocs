

# Create a presentation document by providing a file name (Open XML SDK)

This topic shows how to use the classes in the Open XML SDK 2.5 to
create a presentation document programmatically.

To use the sample code in this topic, you must install the [Open XML SDK 2.5](https://www.nuget.org/packages/DocumentFormat.OpenXml/2.5.0). You
must explicitly reference the following assemblies in your project:

- WindowsBase

- DocumentFormat.OpenXml (Installed by the Open XML SDK)

You must also use the following **using**
directives or **Imports** statements to compile
the code in this topic.

```csharp
    using DocumentFormat.OpenXml;
    using DocumentFormat.OpenXml.Drawing;
    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.Presentation;
    using P = DocumentFormat.OpenXml.Presentation;
    using D = DocumentFormat.OpenXml.Drawing;
```



--------------------------------------------------------------------------------

## Create a Presentation

A presentation file, like all files defined by the Open XML standard,
consists of a package file container. This is the file that users see in
their file explorer; it usually has a .pptx extension. The package file
is represented in the Open XML SDK 2.5 by the [PresentationDocument](https://msdn.microsoft.com/library/office/documentformat.openxml.packaging.presentationdocument.aspx) class. The
presentation document contains, among other parts, a presentation part.
The presentation part, represented in the Open XML SDK 2.5 by the [PresentationPart](https://msdn.microsoft.com/library/office/documentformat.openxml.packaging.presentationpart.aspx) class, contains the basic
*PresentationML* definition for the slide presentation. PresentationML
is the markup language used for creating presentations. Each package can
contain only one presentation part, and its root element must be
\<presentation\>.

The API calls used to create a new presentation document package are
relatively simple. The first step is to call the static [Create(String,PresentationDocumentType)](https://msdn.microsoft.com/library/office/cc535977.aspx)
method of the [PresentationDocument](https://msdn.microsoft.com/library/office/documentformat.openxml.packaging.presentationdocument.aspx) class, as shown here
in the **CreatePresentation** procedure, which is the first part of the
complete code sample presented later in the article. The
**CreatePresentation** code calls the override of the **Create** method that takes as arguments the path to
the new document and the type of presentation document to be created.
The types of presentation documents available in that argument are
defined by a [PresentationDocumentType](https://msdn.microsoft.com/library/office/documentformat.openxml.presentationdocumenttype.aspx) enumerated value.

Next, the code calls [AddPresentationPart()](https://msdn.microsoft.com/library/office/documentformat.openxml.packaging.presentationdocument.addpresentationpart.aspx), which creates and
returns a **PresentationPart**. After the **PresentationPart** class instance is created, a new
root element for the presentation is added by setting the [Presentation](https://msdn.microsoft.com/library/office/documentformat.openxml.packaging.presentationpart.presentation.aspx) property equal to the instance
of the [Presentation](https://msdn.microsoft.com/library/office/documentformat.openxml.presentation.presentation.aspx) class returned from a call to
the **Presentation** class constructor.

In order to create a complete, useable, and valid presentation, the code
must also add a number of other parts to the presentation package. In
the example code, this is taken care of by a call to a utility function
named **CreatePresentationsParts**. That function then calls a number of
other utility functions that, taken together, create all the
presentation parts needed for a basic presentation, including slide,
slide layout, slide master, and theme parts.

```csharp
    public static void CreatePresentation(string filepath)
    {
        // Create a presentation at a specified file path. The presentation document type is pptx, by default.
        PresentationDocument presentationDoc = PresentationDocument.Create(filepath, PresentationDocumentType.Presentation);
        PresentationPart presentationPart = presentationDoc.AddPresentationPart();
        presentationPart.Presentation = new Presentation();

        CreatePresentationParts(presentationPart);

        // Close the presentation handle
        presentationDoc.Close();
    }
```

Using the Open XML SDK 2.5, you can create presentation structure and
content by using strongly-typed classes that correspond to
PresentationML elements. You can find these classes in the [DocumentFormat.OpenXml.Presentation](https://msdn.microsoft.com/library/office/documentformat.openxml.presentation.aspx)
namespace. The following table lists the names of the classes that
correspond to the presentation, slide, slide master, slide layout, and
theme elements. The class that corresponds to the theme element is
actually part of the [DocumentFormat.OpenXml.Drawing](https://msdn.microsoft.com/library/office/documentformat.openxml.drawing.aspx) namespace.
Themes are common to all Open XML markup languages.

| PresentationML Element | Open XML SDK 2.5 Class |
|---|---|
| &lt;presentation&gt; | [Presentation](https://msdn.microsoft.com/library/office/documentformat.openxml.presentation.presentation.aspx) |
| &lt;sld&gt; | [Slide](https://msdn.microsoft.com/library/office/documentformat.openxml.presentation.slide.aspx) |
| &lt;sldMaster&gt; | [SlideMaster](https://msdn.microsoft.com/library/office/documentformat.openxml.presentation.slidemaster.aspx) |
| &lt;sldLayout&gt; | [SlideLayout](https://msdn.microsoft.com/library/office/documentformat.openxml.presentation.slidelayout.aspx) |
| &lt;theme&gt; | [Theme](https://msdn.microsoft.com/library/office/documentformat.openxml.drawing.theme.aspx) |

The PresentationML code that follows is the XML in the presentation part
(in the file presentation.xml) for a simple presentation that contains
two slides.

```xml
    <p:presentation xmlns:p="…" … >
      <p:sldMasterIdLst>
        <p:sldMasterId xmlns:rel="https://…/relationships" rel:id="rId1"/>
      </p:sldMasterIdLst>
      <p:notesMasterIdLst>
        <p:notesMasterId xmlns:rel="https://…/relationships" rel:id="rId4"/>
      </p:notesMasterIdLst>
      <p:handoutMasterIdLst>
        <p:handoutMasterId xmlns:rel="https://…/relationships" rel:id="rId5"/>
      </p:handoutMasterIdLst>
      <p:sldIdLst>
        <p:sldId id="267" xmlns:rel="https://…/relationships" rel:id="rId2"/>
        <p:sldId id="256" xmlns:rel="https://…/relationships" rel:id="rId3"/>
      </p:sldIdLst>
      <p:sldSz cx="9144000" cy="6858000"/>
      <p:notesSz cx="6858000" cy="9144000"/>
    </p:presentation>
```

--------------------------------------------------------------------------------

## Sample Code

Following is the complete sample C\# and VB code to create a
presentation, given a file path.

```csharp
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;
    using DocumentFormat.OpenXml;
    using DocumentFormat.OpenXml.Drawing;
    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.Presentation;
    using P = DocumentFormat.OpenXml.Presentation;
    using D = DocumentFormat.OpenXml.Drawing;

    namespace CreatePresentationDocument
    {
        class Program
        {
            static void Main(string[] args)
            {
                string filepath = @"C:\Users\username\Documents\PresentationFromFilename.pptx";
                CreatePresentation(filepath);
            } 

            public static void CreatePresentation(string filepath)
            {
                // Create a presentation at a specified file path. The presentation document type is pptx, by default.
                PresentationDocument presentationDoc = PresentationDocument.Create(filepath, PresentationDocumentType.Presentation);
                PresentationPart presentationPart = presentationDoc.AddPresentationPart();
                presentationPart.Presentation = new Presentation();

                CreatePresentationParts(presentationPart);            

                //Close the presentation handle
                presentationDoc.Close();
            } 

            private static void CreatePresentationParts(PresentationPart presentationPart)
            {
                SlideMasterIdList slideMasterIdList1 = new SlideMasterIdList(new SlideMasterId() { Id = (UInt32Value)2147483648U, RelationshipId = "rId1" });
                SlideIdList slideIdList1 = new SlideIdList(new SlideId() { Id = (UInt32Value)256U, RelationshipId = "rId2" });
                SlideSize slideSize1 = new SlideSize() { Cx = 9144000, Cy = 6858000, Type = SlideSizeValues.Screen4x3 };
                NotesSize notesSize1 = new NotesSize() { Cx = 6858000, Cy = 9144000 };
                DefaultTextStyle defaultTextStyle1 = new DefaultTextStyle();

               presentationPart.Presentation.Append(slideMasterIdList1, slideIdList1, slideSize1, notesSize1, defaultTextStyle1);

               SlidePart slidePart1;
               SlideLayoutPart slideLayoutPart1;
               SlideMasterPart slideMasterPart1;
               ThemePart themePart1;

                
                slidePart1 = CreateSlidePart(presentationPart);
                slideLayoutPart1 = CreateSlideLayoutPart(slidePart1);
                slideMasterPart1 = CreateSlideMasterPart(slideLayoutPart1);
                themePart1 = CreateTheme(slideMasterPart1); 
      
                slideMasterPart1.AddPart(slideLayoutPart1, "rId1");
                presentationPart.AddPart(slideMasterPart1, "rId1");
                presentationPart.AddPart(themePart1, "rId5");            
            }

        private static SlidePart CreateSlidePart(PresentationPart presentationPart)        
            {
                SlidePart slidePart1 = presentationPart.AddNewPart<SlidePart>("rId2");
                    slidePart1.Slide = new Slide(
                            new CommonSlideData(
                                new ShapeTree(
                                    new P.NonVisualGroupShapeProperties(
                                        new P.NonVisualDrawingProperties() { Id = (UInt32Value)1U, Name = "" },
                                        new P.NonVisualGroupShapeDrawingProperties(),
                                        new ApplicationNonVisualDrawingProperties()),
                                    new GroupShapeProperties(new TransformGroup()),
                                    new P.Shape(
                                        new P.NonVisualShapeProperties(
                                            new P.NonVisualDrawingProperties() { Id = (UInt32Value)2U, Name = "Title 1" },
                                            new P.NonVisualShapeDrawingProperties(new ShapeLocks() { NoGrouping = true }),
                                            new ApplicationNonVisualDrawingProperties(new PlaceholderShape())),
                                        new P.ShapeProperties(),
                                        new P.TextBody(
                                            new BodyProperties(),
                                            new ListStyle(),
                                            new Paragraph(new EndParagraphRunProperties() { Language = "en-US" }))))),
                            new ColorMapOverride(new MasterColorMapping()));
                    return slidePart1;
             } 
       
          private static SlideLayoutPart CreateSlideLayoutPart(SlidePart slidePart1)
            {
                SlideLayoutPart slideLayoutPart1 = slidePart1.AddNewPart<SlideLayoutPart>("rId1");
                SlideLayout slideLayout = new SlideLayout(
                new CommonSlideData(new ShapeTree(
                  new P.NonVisualGroupShapeProperties(
                  new P.NonVisualDrawingProperties() { Id = (UInt32Value)1U, Name = "" },
                  new P.NonVisualGroupShapeDrawingProperties(),
                  new ApplicationNonVisualDrawingProperties()),
                  new GroupShapeProperties(new TransformGroup()),
                  new P.Shape(
                  new P.NonVisualShapeProperties(
                    new P.NonVisualDrawingProperties() { Id = (UInt32Value)2U, Name = "" },
                    new P.NonVisualShapeDrawingProperties(new ShapeLocks() { NoGrouping = true }),
                    new ApplicationNonVisualDrawingProperties(new PlaceholderShape())),
                  new P.ShapeProperties(),
                  new P.TextBody(
                    new BodyProperties(),
                    new ListStyle(),
                    new Paragraph(new EndParagraphRunProperties()))))),
                new ColorMapOverride(new MasterColorMapping()));
                slideLayoutPart1.SlideLayout = slideLayout;
                return slideLayoutPart1;
             }

       private static SlideMasterPart CreateSlideMasterPart(SlideLayoutPart slideLayoutPart1)
       {
           SlideMasterPart slideMasterPart1 = slideLayoutPart1.AddNewPart<SlideMasterPart>("rId1");
           SlideMaster slideMaster = new SlideMaster(
           new CommonSlideData(new ShapeTree(
             new P.NonVisualGroupShapeProperties(
             new P.NonVisualDrawingProperties() { Id = (UInt32Value)1U, Name = "" },
             new P.NonVisualGroupShapeDrawingProperties(),
             new ApplicationNonVisualDrawingProperties()),
             new GroupShapeProperties(new TransformGroup()),
             new P.Shape(
             new P.NonVisualShapeProperties(
               new P.NonVisualDrawingProperties() { Id = (UInt32Value)2U, Name = "Title Placeholder 1" },
               new P.NonVisualShapeDrawingProperties(new ShapeLocks() { NoGrouping = true }),
               new ApplicationNonVisualDrawingProperties(new PlaceholderShape() { Type = PlaceholderValues.Title })),
             new P.ShapeProperties(),
             new P.TextBody(
               new BodyProperties(),
               new ListStyle(),
               new Paragraph())))),
           new P.ColorMap() { Background1 = D.ColorSchemeIndexValues.Light1, Text1 = D.ColorSchemeIndexValues.Dark1, Background2 = D.ColorSchemeIndexValues.Light2, Text2 = D.ColorSchemeIndexValues.Dark2, Accent1 = D.ColorSchemeIndexValues.Accent1, Accent2 = D.ColorSchemeIndexValues.Accent2, Accent3 = D.ColorSchemeIndexValues.Accent3, Accent4 = D.ColorSchemeIndexValues.Accent4, Accent5 = D.ColorSchemeIndexValues.Accent5, Accent6 = D.ColorSchemeIndexValues.Accent6, Hyperlink = D.ColorSchemeIndexValues.Hyperlink, FollowedHyperlink = D.ColorSchemeIndexValues.FollowedHyperlink },
           new SlideLayoutIdList(new SlideLayoutId() { Id = (UInt32Value)2147483649U, RelationshipId = "rId1" }),
           new TextStyles(new TitleStyle(), new BodyStyle(), new OtherStyle()));
           slideMasterPart1.SlideMaster = slideMaster;

           return slideMasterPart1;
        }

       private static ThemePart CreateTheme(SlideMasterPart slideMasterPart1)
       {
           ThemePart themePart1 = slideMasterPart1.AddNewPart<ThemePart>("rId5");
           D.Theme theme1 = new D.Theme() { Name = "Office Theme" };

           D.ThemeElements themeElements1 = new D.ThemeElements(
           new D.ColorScheme(
             new D.Dark1Color(new D.SystemColor() { Val = D.SystemColorValues.WindowText, LastColor = "000000" }),
             new D.Light1Color(new D.SystemColor() { Val = D.SystemColorValues.Window, LastColor = "FFFFFF" }),
             new D.Dark2Color(new D.RgbColorModelHex() { Val = "1F497D" }),
             new D.Light2Color(new D.RgbColorModelHex() { Val = "EEECE1" }),
             new D.Accent1Color(new D.RgbColorModelHex() { Val = "4F81BD" }),
             new D.Accent2Color(new D.RgbColorModelHex() { Val = "C0504D" }),
             new D.Accent3Color(new D.RgbColorModelHex() { Val = "9BBB59" }),
             new D.Accent4Color(new D.RgbColorModelHex() { Val = "8064A2" }),
             new D.Accent5Color(new D.RgbColorModelHex() { Val = "4BACC6" }),
             new D.Accent6Color(new D.RgbColorModelHex() { Val = "F79646" }),
             new D.Hyperlink(new D.RgbColorModelHex() { Val = "0000FF" }),
             new D.FollowedHyperlinkColor(new D.RgbColorModelHex() { Val = "800080" })) { Name = "Office" },
             new D.FontScheme(
             new D.MajorFont(
             new D.LatinFont() { Typeface = "Calibri" },
             new D.EastAsianFont() { Typeface = "" },
             new D.ComplexScriptFont() { Typeface = "" }),
             new D.MinorFont(
             new D.LatinFont() { Typeface = "Calibri" },
             new D.EastAsianFont() { Typeface = "" },
             new D.ComplexScriptFont() { Typeface = "" })) { Name = "Office" },
             new D.FormatScheme(
             new D.FillStyleList(
             new D.SolidFill(new D.SchemeColor() { Val = D.SchemeColorValues.PhColor }),
             new D.GradientFill(
               new D.GradientStopList(
               new D.GradientStop(new D.SchemeColor(new D.Tint() { Val = 50000 },
                 new D.SaturationModulation() { Val = 300000 }) { Val = D.SchemeColorValues.PhColor }) { Position = 0 },
               new D.GradientStop(new D.SchemeColor(new D.Tint() { Val = 37000 },
                new D.SaturationModulation() { Val = 300000 }) { Val = D.SchemeColorValues.PhColor }) { Position = 35000 },
               new D.GradientStop(new D.SchemeColor(new D.Tint() { Val = 15000 },
                new D.SaturationModulation() { Val = 350000 }) { Val = D.SchemeColorValues.PhColor }) { Position = 100000 }
               ),
               new D.LinearGradientFill() { Angle = 16200000, Scaled = true }),
             new D.NoFill(),
             new D.PatternFill(),
             new D.GroupFill()),
             new D.LineStyleList(
             new D.Outline(
               new D.SolidFill(
               new D.SchemeColor(
                 new D.Shade() { Val = 95000 },
                 new D.SaturationModulation() { Val = 105000 }) { Val = D.SchemeColorValues.PhColor }),
               new D.PresetDash() { Val = D.PresetLineDashValues.Solid })
             {
                 Width = 9525,
                 CapType = D.LineCapValues.Flat,
                 CompoundLineType = D.CompoundLineValues.Single,
                 Alignment = D.PenAlignmentValues.Center
             },
             new D.Outline(
               new D.SolidFill(
               new D.SchemeColor(
                 new D.Shade() { Val = 95000 },
                 new D.SaturationModulation() { Val = 105000 }) { Val = D.SchemeColorValues.PhColor }),
               new D.PresetDash() { Val = D.PresetLineDashValues.Solid })
             {
                 Width = 9525,
                 CapType = D.LineCapValues.Flat,
                 CompoundLineType = D.CompoundLineValues.Single,
                 Alignment = D.PenAlignmentValues.Center
             },
             new D.Outline(
               new D.SolidFill(
               new D.SchemeColor(
                 new D.Shade() { Val = 95000 },
                 new D.SaturationModulation() { Val = 105000 }) { Val = D.SchemeColorValues.PhColor }),
               new D.PresetDash() { Val = D.PresetLineDashValues.Solid })
             {
                 Width = 9525,
                 CapType = D.LineCapValues.Flat,
                 CompoundLineType = D.CompoundLineValues.Single,
                 Alignment = D.PenAlignmentValues.Center
             }),
             new D.EffectStyleList(
             new D.EffectStyle(
               new D.EffectList(
               new D.OuterShadow(
                 new D.RgbColorModelHex(
                 new D.Alpha() { Val = 38000 }) { Val = "000000" }) { BlurRadius = 40000L, Distance = 20000L, Direction = 5400000, RotateWithShape = false })),
             new D.EffectStyle(
               new D.EffectList(
               new D.OuterShadow(
                 new D.RgbColorModelHex(
                 new D.Alpha() { Val = 38000 }) { Val = "000000" }) { BlurRadius = 40000L, Distance = 20000L, Direction = 5400000, RotateWithShape = false })),
             new D.EffectStyle(
               new D.EffectList(
               new D.OuterShadow(
                 new D.RgbColorModelHex(
                 new D.Alpha() { Val = 38000 }) { Val = "000000" }) { BlurRadius = 40000L, Distance = 20000L, Direction = 5400000, RotateWithShape = false }))),
             new D.BackgroundFillStyleList(
             new D.SolidFill(new D.SchemeColor() { Val = D.SchemeColorValues.PhColor }),
             new D.GradientFill(
               new D.GradientStopList(
               new D.GradientStop(
                 new D.SchemeColor(new D.Tint() { Val = 50000 },
                   new D.SaturationModulation() { Val = 300000 }) { Val = D.SchemeColorValues.PhColor }) { Position = 0 },
               new D.GradientStop(
                 new D.SchemeColor(new D.Tint() { Val = 50000 },
                   new D.SaturationModulation() { Val = 300000 }) { Val = D.SchemeColorValues.PhColor }) { Position = 0 },
               new D.GradientStop(
                 new D.SchemeColor(new D.Tint() { Val = 50000 },
                   new D.SaturationModulation() { Val = 300000 }) { Val = D.SchemeColorValues.PhColor }) { Position = 0 }),
               new D.LinearGradientFill() { Angle = 16200000, Scaled = true }),
             new D.GradientFill(
               new D.GradientStopList(
               new D.GradientStop(
                 new D.SchemeColor(new D.Tint() { Val = 50000 },
                   new D.SaturationModulation() { Val = 300000 }) { Val = D.SchemeColorValues.PhColor }) { Position = 0 },
               new D.GradientStop(
                 new D.SchemeColor(new D.Tint() { Val = 50000 },
                   new D.SaturationModulation() { Val = 300000 }) { Val = D.SchemeColorValues.PhColor }) { Position = 0 }),
               new D.LinearGradientFill() { Angle = 16200000, Scaled = true }))) { Name = "Office" });

           theme1.Append(themeElements1);
           theme1.Append(new D.ObjectDefaults());
           theme1.Append(new D.ExtraColorSchemeList());

           themePart1.Theme = theme1;
           return themePart1;

             }
        } 
    } 
```



--------------------------------------------------------------------------------

## See also 

[About the Open XML SDK 2.5 for Office](about-the-open-xml-sdk.md)  

[Structure of a PresentationML Document](structure-of-a-presentationml-document.md)  

[How to: Insert a new slide into a presentation (Open XML SDK)](how-to-insert-a-new-slide-into-a-presentation.md)  

[How to: Delete a slide from a presentation (Open XML SDK)](how-to-delete-a-slide-from-a-presentation.md)  

[How to: Retrieve the number of slides in a presentation document (Open XML SDK)](how-to-retrieve-the-number-of-slides-in-a-presentation-document.md)  

[How to: Apply a theme to a presentation (Open XML SDK)](how-to-apply-a-theme-to-a-presentation.md)  

- [Open XML SDK 2.5 class library reference](/office/open-xml/open-xml-sdk)

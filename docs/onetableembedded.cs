using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml;
using A = DocumentFormat.OpenXml.Drawing;
using P14 = DocumentFormat.OpenXml.Office2010.PowerPoint;
using P15 = DocumentFormat.OpenXml.Office2013.PowerPoint;
using Thm15 = DocumentFormat.OpenXml.Office2013.Theme;
using Ap = DocumentFormat.OpenXml.ExtendedProperties;
using Vt = DocumentFormat.OpenXml.VariantTypes;

namespace GeneratedCode
{
    public class GeneratedClass
    {
        // Creates a PresentationDocument.
        public void CreatePackage(string filePath)
        {
            using(PresentationDocument package = PresentationDocument.Create(filePath, PresentationDocumentType.Presentation))
            {
                CreateParts(package);
            }
        }

        // Adds child parts and generates content of the specified part.
        private void CreateParts(PresentationDocument document)
        {
            ThumbnailPart thumbnailPart1 = document.AddNewPart<ThumbnailPart>("image/jpeg", "rId2");
            GenerateThumbnailPart1Content(thumbnailPart1);

            PresentationPart presentationPart1 = document.AddPresentationPart();
            GeneratePresentationPart1Content(presentationPart1);

            PresentationPropertiesPart presentationPropertiesPart1 = presentationPart1.AddNewPart<PresentationPropertiesPart>("rId3");
            GeneratePresentationPropertiesPart1Content(presentationPropertiesPart1);

            SlidePart slidePart1 = presentationPart1.AddNewPart<SlidePart>("rId2");
            GenerateSlidePart1Content(slidePart1);

            ImagePart imagePart1 = slidePart1.AddNewPart<ImagePart>("image/x-emf", "rId3");
            GenerateImagePart1Content(imagePart1);

            SlideLayoutPart slideLayoutPart1 = slidePart1.AddNewPart<SlideLayoutPart>("rId1");
            GenerateSlideLayoutPart1Content(slideLayoutPart1);

            SlideMasterPart slideMasterPart1 = slideLayoutPart1.AddNewPart<SlideMasterPart>("rId1");
            GenerateSlideMasterPart1Content(slideMasterPart1);

            SlideLayoutPart slideLayoutPart2 = slideMasterPart1.AddNewPart<SlideLayoutPart>("rId8");
            GenerateSlideLayoutPart2Content(slideLayoutPart2);

            slideLayoutPart2.AddPart(slideMasterPart1, "rId1");

            SlideLayoutPart slideLayoutPart3 = slideMasterPart1.AddNewPart<SlideLayoutPart>("rId3");
            GenerateSlideLayoutPart3Content(slideLayoutPart3);

            slideLayoutPart3.AddPart(slideMasterPart1, "rId1");

            SlideLayoutPart slideLayoutPart4 = slideMasterPart1.AddNewPart<SlideLayoutPart>("rId7");
            GenerateSlideLayoutPart4Content(slideLayoutPart4);

            slideLayoutPart4.AddPart(slideMasterPart1, "rId1");

            ThemePart themePart1 = slideMasterPart1.AddNewPart<ThemePart>("rId12");
            GenerateThemePart1Content(themePart1);

            SlideLayoutPart slideLayoutPart5 = slideMasterPart1.AddNewPart<SlideLayoutPart>("rId2");
            GenerateSlideLayoutPart5Content(slideLayoutPart5);

            slideLayoutPart5.AddPart(slideMasterPart1, "rId1");

            slideMasterPart1.AddPart(slideLayoutPart1, "rId1");

            SlideLayoutPart slideLayoutPart6 = slideMasterPart1.AddNewPart<SlideLayoutPart>("rId6");
            GenerateSlideLayoutPart6Content(slideLayoutPart6);

            slideLayoutPart6.AddPart(slideMasterPart1, "rId1");

            SlideLayoutPart slideLayoutPart7 = slideMasterPart1.AddNewPart<SlideLayoutPart>("rId11");
            GenerateSlideLayoutPart7Content(slideLayoutPart7);

            slideLayoutPart7.AddPart(slideMasterPart1, "rId1");

            SlideLayoutPart slideLayoutPart8 = slideMasterPart1.AddNewPart<SlideLayoutPart>("rId5");
            GenerateSlideLayoutPart8Content(slideLayoutPart8);

            slideLayoutPart8.AddPart(slideMasterPart1, "rId1");

            SlideLayoutPart slideLayoutPart9 = slideMasterPart1.AddNewPart<SlideLayoutPart>("rId10");
            GenerateSlideLayoutPart9Content(slideLayoutPart9);

            slideLayoutPart9.AddPart(slideMasterPart1, "rId1");

            SlideLayoutPart slideLayoutPart10 = slideMasterPart1.AddNewPart<SlideLayoutPart>("rId4");
            GenerateSlideLayoutPart10Content(slideLayoutPart10);

            slideLayoutPart10.AddPart(slideMasterPart1, "rId1");

            SlideLayoutPart slideLayoutPart11 = slideMasterPart1.AddNewPart<SlideLayoutPart>("rId9");
            GenerateSlideLayoutPart11Content(slideLayoutPart11);

            slideLayoutPart11.AddPart(slideMasterPart1, "rId1");

            slidePart1.AddExternalRelationship("http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject", new System.Uri("file:///C:\\Users\\matej.kratochvil\\onetable.xlsx!List1!RWATAB", System.UriKind.Absolute), "rId2");
            presentationPart1.AddPart(slideMasterPart1, "rId1");

            TableStylesPart tableStylesPart1 = presentationPart1.AddNewPart<TableStylesPart>("rId6");
            GenerateTableStylesPart1Content(tableStylesPart1);

            presentationPart1.AddPart(themePart1, "rId5");

            ViewPropertiesPart viewPropertiesPart1 = presentationPart1.AddNewPart<ViewPropertiesPart>("rId4");
            GenerateViewPropertiesPart1Content(viewPropertiesPart1);

            ExtendedFilePropertiesPart extendedFilePropertiesPart1 = document.AddNewPart<ExtendedFilePropertiesPart>("rId4");
            GenerateExtendedFilePropertiesPart1Content(extendedFilePropertiesPart1);

            SetPackageProperties(document);
        }

        // Generates content of thumbnailPart1.
        private void GenerateThumbnailPart1Content(ThumbnailPart thumbnailPart1)
        {
            System.IO.Stream data = GetBinaryDataStream(thumbnailPart1Data);
            thumbnailPart1.FeedData(data);
            data.Close();
        }

        // Generates content of presentationPart1.
        private void GeneratePresentationPart1Content(PresentationPart presentationPart1)
        {
            Presentation presentation1 = new Presentation(){ SaveSubsetFonts = true };
            presentation1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            presentation1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            presentation1.AddNamespaceDeclaration("p", "http://schemas.openxmlformats.org/presentationml/2006/main");

            SlideMasterIdList slideMasterIdList1 = new SlideMasterIdList();
            SlideMasterId slideMasterId1 = new SlideMasterId(){ Id = (UInt32Value)2147483660U, RelationshipId = "rId1" };

            slideMasterIdList1.Append(slideMasterId1);

            SlideIdList slideIdList1 = new SlideIdList();
            SlideId slideId1 = new SlideId(){ Id = (UInt32Value)256U, RelationshipId = "rId2" };

            slideIdList1.Append(slideId1);
            SlideSize slideSize1 = new SlideSize(){ Cx = 9144000, Cy = 6858000, Type = SlideSizeValues.Screen4x3 };
            NotesSize notesSize1 = new NotesSize(){ Cx = 6858000L, Cy = 9144000L };

            DefaultTextStyle defaultTextStyle1 = new DefaultTextStyle();

            A.DefaultParagraphProperties defaultParagraphProperties1 = new A.DefaultParagraphProperties();
            A.DefaultRunProperties defaultRunProperties1 = new A.DefaultRunProperties(){ Language = "en-US" };

            defaultParagraphProperties1.Append(defaultRunProperties1);

            A.Level1ParagraphProperties level1ParagraphProperties1 = new A.Level1ParagraphProperties(){ LeftMargin = 0, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 457200, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.DefaultRunProperties defaultRunProperties2 = new A.DefaultRunProperties(){ FontSize = 1800, Kerning = 1200 };

            A.SolidFill solidFill1 = new A.SolidFill();
            A.SchemeColor schemeColor1 = new A.SchemeColor(){ Val = A.SchemeColorValues.Text1 };

            solidFill1.Append(schemeColor1);
            A.LatinFont latinFont1 = new A.LatinFont(){ Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont1 = new A.EastAsianFont(){ Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont1 = new A.ComplexScriptFont(){ Typeface = "+mn-cs" };

            defaultRunProperties2.Append(solidFill1);
            defaultRunProperties2.Append(latinFont1);
            defaultRunProperties2.Append(eastAsianFont1);
            defaultRunProperties2.Append(complexScriptFont1);

            level1ParagraphProperties1.Append(defaultRunProperties2);

            A.Level2ParagraphProperties level2ParagraphProperties1 = new A.Level2ParagraphProperties(){ LeftMargin = 457200, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 457200, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.DefaultRunProperties defaultRunProperties3 = new A.DefaultRunProperties(){ FontSize = 1800, Kerning = 1200 };

            A.SolidFill solidFill2 = new A.SolidFill();
            A.SchemeColor schemeColor2 = new A.SchemeColor(){ Val = A.SchemeColorValues.Text1 };

            solidFill2.Append(schemeColor2);
            A.LatinFont latinFont2 = new A.LatinFont(){ Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont2 = new A.EastAsianFont(){ Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont2 = new A.ComplexScriptFont(){ Typeface = "+mn-cs" };

            defaultRunProperties3.Append(solidFill2);
            defaultRunProperties3.Append(latinFont2);
            defaultRunProperties3.Append(eastAsianFont2);
            defaultRunProperties3.Append(complexScriptFont2);

            level2ParagraphProperties1.Append(defaultRunProperties3);

            A.Level3ParagraphProperties level3ParagraphProperties1 = new A.Level3ParagraphProperties(){ LeftMargin = 914400, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 457200, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.DefaultRunProperties defaultRunProperties4 = new A.DefaultRunProperties(){ FontSize = 1800, Kerning = 1200 };

            A.SolidFill solidFill3 = new A.SolidFill();
            A.SchemeColor schemeColor3 = new A.SchemeColor(){ Val = A.SchemeColorValues.Text1 };

            solidFill3.Append(schemeColor3);
            A.LatinFont latinFont3 = new A.LatinFont(){ Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont3 = new A.EastAsianFont(){ Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont3 = new A.ComplexScriptFont(){ Typeface = "+mn-cs" };

            defaultRunProperties4.Append(solidFill3);
            defaultRunProperties4.Append(latinFont3);
            defaultRunProperties4.Append(eastAsianFont3);
            defaultRunProperties4.Append(complexScriptFont3);

            level3ParagraphProperties1.Append(defaultRunProperties4);

            A.Level4ParagraphProperties level4ParagraphProperties1 = new A.Level4ParagraphProperties(){ LeftMargin = 1371600, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 457200, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.DefaultRunProperties defaultRunProperties5 = new A.DefaultRunProperties(){ FontSize = 1800, Kerning = 1200 };

            A.SolidFill solidFill4 = new A.SolidFill();
            A.SchemeColor schemeColor4 = new A.SchemeColor(){ Val = A.SchemeColorValues.Text1 };

            solidFill4.Append(schemeColor4);
            A.LatinFont latinFont4 = new A.LatinFont(){ Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont4 = new A.EastAsianFont(){ Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont4 = new A.ComplexScriptFont(){ Typeface = "+mn-cs" };

            defaultRunProperties5.Append(solidFill4);
            defaultRunProperties5.Append(latinFont4);
            defaultRunProperties5.Append(eastAsianFont4);
            defaultRunProperties5.Append(complexScriptFont4);

            level4ParagraphProperties1.Append(defaultRunProperties5);

            A.Level5ParagraphProperties level5ParagraphProperties1 = new A.Level5ParagraphProperties(){ LeftMargin = 1828800, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 457200, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.DefaultRunProperties defaultRunProperties6 = new A.DefaultRunProperties(){ FontSize = 1800, Kerning = 1200 };

            A.SolidFill solidFill5 = new A.SolidFill();
            A.SchemeColor schemeColor5 = new A.SchemeColor(){ Val = A.SchemeColorValues.Text1 };

            solidFill5.Append(schemeColor5);
            A.LatinFont latinFont5 = new A.LatinFont(){ Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont5 = new A.EastAsianFont(){ Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont5 = new A.ComplexScriptFont(){ Typeface = "+mn-cs" };

            defaultRunProperties6.Append(solidFill5);
            defaultRunProperties6.Append(latinFont5);
            defaultRunProperties6.Append(eastAsianFont5);
            defaultRunProperties6.Append(complexScriptFont5);

            level5ParagraphProperties1.Append(defaultRunProperties6);

            A.Level6ParagraphProperties level6ParagraphProperties1 = new A.Level6ParagraphProperties(){ LeftMargin = 2286000, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 457200, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.DefaultRunProperties defaultRunProperties7 = new A.DefaultRunProperties(){ FontSize = 1800, Kerning = 1200 };

            A.SolidFill solidFill6 = new A.SolidFill();
            A.SchemeColor schemeColor6 = new A.SchemeColor(){ Val = A.SchemeColorValues.Text1 };

            solidFill6.Append(schemeColor6);
            A.LatinFont latinFont6 = new A.LatinFont(){ Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont6 = new A.EastAsianFont(){ Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont6 = new A.ComplexScriptFont(){ Typeface = "+mn-cs" };

            defaultRunProperties7.Append(solidFill6);
            defaultRunProperties7.Append(latinFont6);
            defaultRunProperties7.Append(eastAsianFont6);
            defaultRunProperties7.Append(complexScriptFont6);

            level6ParagraphProperties1.Append(defaultRunProperties7);

            A.Level7ParagraphProperties level7ParagraphProperties1 = new A.Level7ParagraphProperties(){ LeftMargin = 2743200, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 457200, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.DefaultRunProperties defaultRunProperties8 = new A.DefaultRunProperties(){ FontSize = 1800, Kerning = 1200 };

            A.SolidFill solidFill7 = new A.SolidFill();
            A.SchemeColor schemeColor7 = new A.SchemeColor(){ Val = A.SchemeColorValues.Text1 };

            solidFill7.Append(schemeColor7);
            A.LatinFont latinFont7 = new A.LatinFont(){ Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont7 = new A.EastAsianFont(){ Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont7 = new A.ComplexScriptFont(){ Typeface = "+mn-cs" };

            defaultRunProperties8.Append(solidFill7);
            defaultRunProperties8.Append(latinFont7);
            defaultRunProperties8.Append(eastAsianFont7);
            defaultRunProperties8.Append(complexScriptFont7);

            level7ParagraphProperties1.Append(defaultRunProperties8);

            A.Level8ParagraphProperties level8ParagraphProperties1 = new A.Level8ParagraphProperties(){ LeftMargin = 3200400, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 457200, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.DefaultRunProperties defaultRunProperties9 = new A.DefaultRunProperties(){ FontSize = 1800, Kerning = 1200 };

            A.SolidFill solidFill8 = new A.SolidFill();
            A.SchemeColor schemeColor8 = new A.SchemeColor(){ Val = A.SchemeColorValues.Text1 };

            solidFill8.Append(schemeColor8);
            A.LatinFont latinFont8 = new A.LatinFont(){ Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont8 = new A.EastAsianFont(){ Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont8 = new A.ComplexScriptFont(){ Typeface = "+mn-cs" };

            defaultRunProperties9.Append(solidFill8);
            defaultRunProperties9.Append(latinFont8);
            defaultRunProperties9.Append(eastAsianFont8);
            defaultRunProperties9.Append(complexScriptFont8);

            level8ParagraphProperties1.Append(defaultRunProperties9);

            A.Level9ParagraphProperties level9ParagraphProperties1 = new A.Level9ParagraphProperties(){ LeftMargin = 3657600, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 457200, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.DefaultRunProperties defaultRunProperties10 = new A.DefaultRunProperties(){ FontSize = 1800, Kerning = 1200 };

            A.SolidFill solidFill9 = new A.SolidFill();
            A.SchemeColor schemeColor9 = new A.SchemeColor(){ Val = A.SchemeColorValues.Text1 };

            solidFill9.Append(schemeColor9);
            A.LatinFont latinFont9 = new A.LatinFont(){ Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont9 = new A.EastAsianFont(){ Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont9 = new A.ComplexScriptFont(){ Typeface = "+mn-cs" };

            defaultRunProperties10.Append(solidFill9);
            defaultRunProperties10.Append(latinFont9);
            defaultRunProperties10.Append(eastAsianFont9);
            defaultRunProperties10.Append(complexScriptFont9);

            level9ParagraphProperties1.Append(defaultRunProperties10);

            defaultTextStyle1.Append(defaultParagraphProperties1);
            defaultTextStyle1.Append(level1ParagraphProperties1);
            defaultTextStyle1.Append(level2ParagraphProperties1);
            defaultTextStyle1.Append(level3ParagraphProperties1);
            defaultTextStyle1.Append(level4ParagraphProperties1);
            defaultTextStyle1.Append(level5ParagraphProperties1);
            defaultTextStyle1.Append(level6ParagraphProperties1);
            defaultTextStyle1.Append(level7ParagraphProperties1);
            defaultTextStyle1.Append(level8ParagraphProperties1);
            defaultTextStyle1.Append(level9ParagraphProperties1);

            presentation1.Append(slideMasterIdList1);
            presentation1.Append(slideIdList1);
            presentation1.Append(slideSize1);
            presentation1.Append(notesSize1);
            presentation1.Append(defaultTextStyle1);

            presentationPart1.Presentation = presentation1;
        }

        // Generates content of presentationPropertiesPart1.
        private void GeneratePresentationPropertiesPart1Content(PresentationPropertiesPart presentationPropertiesPart1)
        {
            PresentationProperties presentationProperties1 = new PresentationProperties();
            presentationProperties1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            presentationProperties1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            presentationProperties1.AddNamespaceDeclaration("p", "http://schemas.openxmlformats.org/presentationml/2006/main");

            PresentationPropertiesExtensionList presentationPropertiesExtensionList1 = new PresentationPropertiesExtensionList();

            PresentationPropertiesExtension presentationPropertiesExtension1 = new PresentationPropertiesExtension(){ Uri = "{E76CE94A-603C-4142-B9EB-6D1370010A27}" };

            P14.DiscardImageEditData discardImageEditData1 = new P14.DiscardImageEditData(){ Val = false };
            discardImageEditData1.AddNamespaceDeclaration("p14", "http://schemas.microsoft.com/office/powerpoint/2010/main");

            presentationPropertiesExtension1.Append(discardImageEditData1);

            PresentationPropertiesExtension presentationPropertiesExtension2 = new PresentationPropertiesExtension(){ Uri = "{D31A062A-798A-4329-ABDD-BBA856620510}" };

            P14.DefaultImageDpi defaultImageDpi1 = new P14.DefaultImageDpi(){ Val = (UInt32Value)32767U };
            defaultImageDpi1.AddNamespaceDeclaration("p14", "http://schemas.microsoft.com/office/powerpoint/2010/main");

            presentationPropertiesExtension2.Append(defaultImageDpi1);

            PresentationPropertiesExtension presentationPropertiesExtension3 = new PresentationPropertiesExtension(){ Uri = "{FD5EFAAD-0ECE-453E-9831-46B23BE46B34}" };

            P15.ChartTrackingReferenceBased chartTrackingReferenceBased1 = new P15.ChartTrackingReferenceBased(){ Val = true };
            chartTrackingReferenceBased1.AddNamespaceDeclaration("p15", "http://schemas.microsoft.com/office/powerpoint/2012/main");

            presentationPropertiesExtension3.Append(chartTrackingReferenceBased1);

            presentationPropertiesExtensionList1.Append(presentationPropertiesExtension1);
            presentationPropertiesExtensionList1.Append(presentationPropertiesExtension2);
            presentationPropertiesExtensionList1.Append(presentationPropertiesExtension3);

            presentationProperties1.Append(presentationPropertiesExtensionList1);

            presentationPropertiesPart1.PresentationProperties = presentationProperties1;
        }

        // Generates content of slidePart1.
        private void GenerateSlidePart1Content(SlidePart slidePart1)
        {
            Slide slide1 = new Slide();
            slide1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            slide1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            slide1.AddNamespaceDeclaration("p", "http://schemas.openxmlformats.org/presentationml/2006/main");

            CommonSlideData commonSlideData1 = new CommonSlideData();

            ShapeTree shapeTree1 = new ShapeTree();

            NonVisualGroupShapeProperties nonVisualGroupShapeProperties1 = new NonVisualGroupShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties1 = new NonVisualDrawingProperties(){ Id = (UInt32Value)1U, Name = "" };
            NonVisualGroupShapeDrawingProperties nonVisualGroupShapeDrawingProperties1 = new NonVisualGroupShapeDrawingProperties();
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties1 = new ApplicationNonVisualDrawingProperties();

            nonVisualGroupShapeProperties1.Append(nonVisualDrawingProperties1);
            nonVisualGroupShapeProperties1.Append(nonVisualGroupShapeDrawingProperties1);
            nonVisualGroupShapeProperties1.Append(applicationNonVisualDrawingProperties1);

            GroupShapeProperties groupShapeProperties1 = new GroupShapeProperties();

            A.TransformGroup transformGroup1 = new A.TransformGroup();
            A.Offset offset1 = new A.Offset(){ X = 0L, Y = 0L };
            A.Extents extents1 = new A.Extents(){ Cx = 0L, Cy = 0L };
            A.ChildOffset childOffset1 = new A.ChildOffset(){ X = 0L, Y = 0L };
            A.ChildExtents childExtents1 = new A.ChildExtents(){ Cx = 0L, Cy = 0L };

            transformGroup1.Append(offset1);
            transformGroup1.Append(extents1);
            transformGroup1.Append(childOffset1);
            transformGroup1.Append(childExtents1);

            groupShapeProperties1.Append(transformGroup1);

            GraphicFrame graphicFrame1 = new GraphicFrame();

            NonVisualGraphicFrameProperties nonVisualGraphicFrameProperties1 = new NonVisualGraphicFrameProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties2 = new NonVisualDrawingProperties(){ Id = (UInt32Value)4U, Name = "Objekt 3" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList1 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension1 = new A.NonVisualDrawingPropertiesExtension(){ Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            OpenXmlUnknownElement openXmlUnknownElement1 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{6B9DEAA6-04FE-B16C-8104-5580AA55F5B0}\" />");

            nonVisualDrawingPropertiesExtension1.Append(openXmlUnknownElement1);

            nonVisualDrawingPropertiesExtensionList1.Append(nonVisualDrawingPropertiesExtension1);

            nonVisualDrawingProperties2.Append(nonVisualDrawingPropertiesExtensionList1);

            NonVisualGraphicFrameDrawingProperties nonVisualGraphicFrameDrawingProperties1 = new NonVisualGraphicFrameDrawingProperties();
            A.GraphicFrameLocks graphicFrameLocks1 = new A.GraphicFrameLocks(){ NoChangeAspect = true };

            nonVisualGraphicFrameDrawingProperties1.Append(graphicFrameLocks1);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties2 = new ApplicationNonVisualDrawingProperties();

            ApplicationNonVisualDrawingPropertiesExtensionList applicationNonVisualDrawingPropertiesExtensionList1 = new ApplicationNonVisualDrawingPropertiesExtensionList();

            ApplicationNonVisualDrawingPropertiesExtension applicationNonVisualDrawingPropertiesExtension1 = new ApplicationNonVisualDrawingPropertiesExtension(){ Uri = "{D42A27DB-BD31-4B8C-83A1-F6EECF244321}" };

            P14.ModificationId modificationId1 = new P14.ModificationId(){ Val = (UInt32Value)3871464956U };
            modificationId1.AddNamespaceDeclaration("p14", "http://schemas.microsoft.com/office/powerpoint/2010/main");

            applicationNonVisualDrawingPropertiesExtension1.Append(modificationId1);

            applicationNonVisualDrawingPropertiesExtensionList1.Append(applicationNonVisualDrawingPropertiesExtension1);

            applicationNonVisualDrawingProperties2.Append(applicationNonVisualDrawingPropertiesExtensionList1);

            nonVisualGraphicFrameProperties1.Append(nonVisualDrawingProperties2);
            nonVisualGraphicFrameProperties1.Append(nonVisualGraphicFrameDrawingProperties1);
            nonVisualGraphicFrameProperties1.Append(applicationNonVisualDrawingProperties2);

            Transform transform1 = new Transform();
            A.Offset offset2 = new A.Offset(){ X = 2516623L, Y = 777875L };
            A.Extents extents2 = new A.Extents(){ Cx = 3576202L, Cy = 5543825L };

            transform1.Append(offset2);
            transform1.Append(extents2);

            A.Graphic graphic1 = new A.Graphic();

            A.GraphicData graphicData1 = new A.GraphicData(){ Uri = "http://schemas.openxmlformats.org/presentationml/2006/ole" };

            AlternateContent alternateContent1 = new AlternateContent();
            alternateContent1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");

            AlternateContentChoice alternateContentChoice1 = new AlternateContentChoice(){ Requires = "v" };
            alternateContentChoice1.AddNamespaceDeclaration("v", "urn:schemas-microsoft-com:vml");

            OleObject oleObject1 = new OleObject(){ Name = "Worksheet", Id = "rId2", ImageWidth = 4914810, ImageHeight = 7620000, ProgId = "Excel.Sheet.12" };
            OleObjectLink oleObjectLink1 = new OleObjectLink(){ AutoUpdate = true };

            oleObject1.Append(oleObjectLink1);

            alternateContentChoice1.Append(oleObject1);

            AlternateContentFallback alternateContentFallback1 = new AlternateContentFallback();

            OleObject oleObject2 = new OleObject(){ Name = "Worksheet", Id = "rId2", ImageWidth = 4914810, ImageHeight = 7620000, ProgId = "Excel.Sheet.12" };
            OleObjectLink oleObjectLink2 = new OleObjectLink(){ AutoUpdate = true };

            Picture picture1 = new Picture();

            NonVisualPictureProperties nonVisualPictureProperties1 = new NonVisualPictureProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties3 = new NonVisualDrawingProperties(){ Id = (UInt32Value)0U, Name = "" };
            NonVisualPictureDrawingProperties nonVisualPictureDrawingProperties1 = new NonVisualPictureDrawingProperties();
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties3 = new ApplicationNonVisualDrawingProperties();

            nonVisualPictureProperties1.Append(nonVisualDrawingProperties3);
            nonVisualPictureProperties1.Append(nonVisualPictureDrawingProperties1);
            nonVisualPictureProperties1.Append(applicationNonVisualDrawingProperties3);

            BlipFill blipFill1 = new BlipFill();
            A.Blip blip1 = new A.Blip(){ Embed = "rId3" };

            A.Stretch stretch1 = new A.Stretch();
            A.FillRectangle fillRectangle1 = new A.FillRectangle();

            stretch1.Append(fillRectangle1);

            blipFill1.Append(blip1);
            blipFill1.Append(stretch1);

            ShapeProperties shapeProperties1 = new ShapeProperties();

            A.Transform2D transform2D1 = new A.Transform2D();
            A.Offset offset3 = new A.Offset(){ X = 2516623L, Y = 777875L };
            A.Extents extents3 = new A.Extents(){ Cx = 3576202L, Cy = 5543825L };

            transform2D1.Append(offset3);
            transform2D1.Append(extents3);

            A.PresetGeometry presetGeometry1 = new A.PresetGeometry(){ Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList1 = new A.AdjustValueList();

            presetGeometry1.Append(adjustValueList1);

            shapeProperties1.Append(transform2D1);
            shapeProperties1.Append(presetGeometry1);

            picture1.Append(nonVisualPictureProperties1);
            picture1.Append(blipFill1);
            picture1.Append(shapeProperties1);

            oleObject2.Append(oleObjectLink2);
            oleObject2.Append(picture1);

            alternateContentFallback1.Append(oleObject2);

            alternateContent1.Append(alternateContentChoice1);
            alternateContent1.Append(alternateContentFallback1);

            graphicData1.Append(alternateContent1);

            graphic1.Append(graphicData1);

            graphicFrame1.Append(nonVisualGraphicFrameProperties1);
            graphicFrame1.Append(transform1);
            graphicFrame1.Append(graphic1);

            shapeTree1.Append(nonVisualGroupShapeProperties1);
            shapeTree1.Append(groupShapeProperties1);
            shapeTree1.Append(graphicFrame1);

            CommonSlideDataExtensionList commonSlideDataExtensionList1 = new CommonSlideDataExtensionList();

            CommonSlideDataExtension commonSlideDataExtension1 = new CommonSlideDataExtension(){ Uri = "{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}" };

            P14.CreationId creationId1 = new P14.CreationId(){ Val = (UInt32Value)1021816136U };
            creationId1.AddNamespaceDeclaration("p14", "http://schemas.microsoft.com/office/powerpoint/2010/main");

            commonSlideDataExtension1.Append(creationId1);

            commonSlideDataExtensionList1.Append(commonSlideDataExtension1);

            commonSlideData1.Append(shapeTree1);
            commonSlideData1.Append(commonSlideDataExtensionList1);

            ColorMapOverride colorMapOverride1 = new ColorMapOverride();
            A.MasterColorMapping masterColorMapping1 = new A.MasterColorMapping();

            colorMapOverride1.Append(masterColorMapping1);

            slide1.Append(commonSlideData1);
            slide1.Append(colorMapOverride1);

            slidePart1.Slide = slide1;
        }

        // Generates content of imagePart1.
        private void GenerateImagePart1Content(ImagePart imagePart1)
        {
            System.IO.Stream data = GetBinaryDataStream(imagePart1Data);
            imagePart1.FeedData(data);
            data.Close();
        }

        // Generates content of slideLayoutPart1.
        private void GenerateSlideLayoutPart1Content(SlideLayoutPart slideLayoutPart1)
        {
            SlideLayout slideLayout1 = new SlideLayout(){ Type = SlideLayoutValues.Title, Preserve = true };
            slideLayout1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            slideLayout1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            slideLayout1.AddNamespaceDeclaration("p", "http://schemas.openxmlformats.org/presentationml/2006/main");

            CommonSlideData commonSlideData2 = new CommonSlideData(){ Name = "Úvodní snímek" };

            ShapeTree shapeTree2 = new ShapeTree();

            NonVisualGroupShapeProperties nonVisualGroupShapeProperties2 = new NonVisualGroupShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties4 = new NonVisualDrawingProperties(){ Id = (UInt32Value)1U, Name = "" };
            NonVisualGroupShapeDrawingProperties nonVisualGroupShapeDrawingProperties2 = new NonVisualGroupShapeDrawingProperties();
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties4 = new ApplicationNonVisualDrawingProperties();

            nonVisualGroupShapeProperties2.Append(nonVisualDrawingProperties4);
            nonVisualGroupShapeProperties2.Append(nonVisualGroupShapeDrawingProperties2);
            nonVisualGroupShapeProperties2.Append(applicationNonVisualDrawingProperties4);

            GroupShapeProperties groupShapeProperties2 = new GroupShapeProperties();

            A.TransformGroup transformGroup2 = new A.TransformGroup();
            A.Offset offset4 = new A.Offset(){ X = 0L, Y = 0L };
            A.Extents extents4 = new A.Extents(){ Cx = 0L, Cy = 0L };
            A.ChildOffset childOffset2 = new A.ChildOffset(){ X = 0L, Y = 0L };
            A.ChildExtents childExtents2 = new A.ChildExtents(){ Cx = 0L, Cy = 0L };

            transformGroup2.Append(offset4);
            transformGroup2.Append(extents4);
            transformGroup2.Append(childOffset2);
            transformGroup2.Append(childExtents2);

            groupShapeProperties2.Append(transformGroup2);

            Shape shape1 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties1 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties5 = new NonVisualDrawingProperties(){ Id = (UInt32Value)2U, Name = "Title 1" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties1 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks1 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties1.Append(shapeLocks1);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties5 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape1 = new PlaceholderShape(){ Type = PlaceholderValues.CenteredTitle };

            applicationNonVisualDrawingProperties5.Append(placeholderShape1);

            nonVisualShapeProperties1.Append(nonVisualDrawingProperties5);
            nonVisualShapeProperties1.Append(nonVisualShapeDrawingProperties1);
            nonVisualShapeProperties1.Append(applicationNonVisualDrawingProperties5);

            ShapeProperties shapeProperties2 = new ShapeProperties();

            A.Transform2D transform2D2 = new A.Transform2D();
            A.Offset offset5 = new A.Offset(){ X = 685800L, Y = 1122363L };
            A.Extents extents5 = new A.Extents(){ Cx = 7772400L, Cy = 2387600L };

            transform2D2.Append(offset5);
            transform2D2.Append(extents5);

            shapeProperties2.Append(transform2D2);

            TextBody textBody1 = new TextBody();
            A.BodyProperties bodyProperties1 = new A.BodyProperties(){ Anchor = A.TextAnchoringTypeValues.Bottom };

            A.ListStyle listStyle1 = new A.ListStyle();

            A.Level1ParagraphProperties level1ParagraphProperties2 = new A.Level1ParagraphProperties(){ Alignment = A.TextAlignmentTypeValues.Center };
            A.DefaultRunProperties defaultRunProperties11 = new A.DefaultRunProperties(){ FontSize = 6000 };

            level1ParagraphProperties2.Append(defaultRunProperties11);

            listStyle1.Append(level1ParagraphProperties2);

            A.Paragraph paragraph1 = new A.Paragraph();

            A.Run run1 = new A.Run();
            A.RunProperties runProperties1 = new A.RunProperties(){ Language = "cs-CZ" };
            A.Text text1 = new A.Text();
            text1.Text = "Kliknutím lze upravit styl.";

            run1.Append(runProperties1);
            run1.Append(text1);
            A.EndParagraphRunProperties endParagraphRunProperties1 = new A.EndParagraphRunProperties(){ Language = "en-US", Dirty = false };

            paragraph1.Append(run1);
            paragraph1.Append(endParagraphRunProperties1);

            textBody1.Append(bodyProperties1);
            textBody1.Append(listStyle1);
            textBody1.Append(paragraph1);

            shape1.Append(nonVisualShapeProperties1);
            shape1.Append(shapeProperties2);
            shape1.Append(textBody1);

            Shape shape2 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties2 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties6 = new NonVisualDrawingProperties(){ Id = (UInt32Value)3U, Name = "Subtitle 2" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties2 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks2 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties2.Append(shapeLocks2);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties6 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape2 = new PlaceholderShape(){ Type = PlaceholderValues.SubTitle, Index = (UInt32Value)1U };

            applicationNonVisualDrawingProperties6.Append(placeholderShape2);

            nonVisualShapeProperties2.Append(nonVisualDrawingProperties6);
            nonVisualShapeProperties2.Append(nonVisualShapeDrawingProperties2);
            nonVisualShapeProperties2.Append(applicationNonVisualDrawingProperties6);

            ShapeProperties shapeProperties3 = new ShapeProperties();

            A.Transform2D transform2D3 = new A.Transform2D();
            A.Offset offset6 = new A.Offset(){ X = 1143000L, Y = 3602038L };
            A.Extents extents6 = new A.Extents(){ Cx = 6858000L, Cy = 1655762L };

            transform2D3.Append(offset6);
            transform2D3.Append(extents6);

            shapeProperties3.Append(transform2D3);

            TextBody textBody2 = new TextBody();
            A.BodyProperties bodyProperties2 = new A.BodyProperties();

            A.ListStyle listStyle2 = new A.ListStyle();

            A.Level1ParagraphProperties level1ParagraphProperties3 = new A.Level1ParagraphProperties(){ LeftMargin = 0, Indent = 0, Alignment = A.TextAlignmentTypeValues.Center };
            A.NoBullet noBullet1 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties12 = new A.DefaultRunProperties(){ FontSize = 2400 };

            level1ParagraphProperties3.Append(noBullet1);
            level1ParagraphProperties3.Append(defaultRunProperties12);

            A.Level2ParagraphProperties level2ParagraphProperties2 = new A.Level2ParagraphProperties(){ LeftMargin = 457200, Indent = 0, Alignment = A.TextAlignmentTypeValues.Center };
            A.NoBullet noBullet2 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties13 = new A.DefaultRunProperties(){ FontSize = 2000 };

            level2ParagraphProperties2.Append(noBullet2);
            level2ParagraphProperties2.Append(defaultRunProperties13);

            A.Level3ParagraphProperties level3ParagraphProperties2 = new A.Level3ParagraphProperties(){ LeftMargin = 914400, Indent = 0, Alignment = A.TextAlignmentTypeValues.Center };
            A.NoBullet noBullet3 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties14 = new A.DefaultRunProperties(){ FontSize = 1800 };

            level3ParagraphProperties2.Append(noBullet3);
            level3ParagraphProperties2.Append(defaultRunProperties14);

            A.Level4ParagraphProperties level4ParagraphProperties2 = new A.Level4ParagraphProperties(){ LeftMargin = 1371600, Indent = 0, Alignment = A.TextAlignmentTypeValues.Center };
            A.NoBullet noBullet4 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties15 = new A.DefaultRunProperties(){ FontSize = 1600 };

            level4ParagraphProperties2.Append(noBullet4);
            level4ParagraphProperties2.Append(defaultRunProperties15);

            A.Level5ParagraphProperties level5ParagraphProperties2 = new A.Level5ParagraphProperties(){ LeftMargin = 1828800, Indent = 0, Alignment = A.TextAlignmentTypeValues.Center };
            A.NoBullet noBullet5 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties16 = new A.DefaultRunProperties(){ FontSize = 1600 };

            level5ParagraphProperties2.Append(noBullet5);
            level5ParagraphProperties2.Append(defaultRunProperties16);

            A.Level6ParagraphProperties level6ParagraphProperties2 = new A.Level6ParagraphProperties(){ LeftMargin = 2286000, Indent = 0, Alignment = A.TextAlignmentTypeValues.Center };
            A.NoBullet noBullet6 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties17 = new A.DefaultRunProperties(){ FontSize = 1600 };

            level6ParagraphProperties2.Append(noBullet6);
            level6ParagraphProperties2.Append(defaultRunProperties17);

            A.Level7ParagraphProperties level7ParagraphProperties2 = new A.Level7ParagraphProperties(){ LeftMargin = 2743200, Indent = 0, Alignment = A.TextAlignmentTypeValues.Center };
            A.NoBullet noBullet7 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties18 = new A.DefaultRunProperties(){ FontSize = 1600 };

            level7ParagraphProperties2.Append(noBullet7);
            level7ParagraphProperties2.Append(defaultRunProperties18);

            A.Level8ParagraphProperties level8ParagraphProperties2 = new A.Level8ParagraphProperties(){ LeftMargin = 3200400, Indent = 0, Alignment = A.TextAlignmentTypeValues.Center };
            A.NoBullet noBullet8 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties19 = new A.DefaultRunProperties(){ FontSize = 1600 };

            level8ParagraphProperties2.Append(noBullet8);
            level8ParagraphProperties2.Append(defaultRunProperties19);

            A.Level9ParagraphProperties level9ParagraphProperties2 = new A.Level9ParagraphProperties(){ LeftMargin = 3657600, Indent = 0, Alignment = A.TextAlignmentTypeValues.Center };
            A.NoBullet noBullet9 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties20 = new A.DefaultRunProperties(){ FontSize = 1600 };

            level9ParagraphProperties2.Append(noBullet9);
            level9ParagraphProperties2.Append(defaultRunProperties20);

            listStyle2.Append(level1ParagraphProperties3);
            listStyle2.Append(level2ParagraphProperties2);
            listStyle2.Append(level3ParagraphProperties2);
            listStyle2.Append(level4ParagraphProperties2);
            listStyle2.Append(level5ParagraphProperties2);
            listStyle2.Append(level6ParagraphProperties2);
            listStyle2.Append(level7ParagraphProperties2);
            listStyle2.Append(level8ParagraphProperties2);
            listStyle2.Append(level9ParagraphProperties2);

            A.Paragraph paragraph2 = new A.Paragraph();

            A.Run run2 = new A.Run();
            A.RunProperties runProperties2 = new A.RunProperties(){ Language = "cs-CZ" };
            A.Text text2 = new A.Text();
            text2.Text = "Kliknutím můžete upravit styl předlohy.";

            run2.Append(runProperties2);
            run2.Append(text2);
            A.EndParagraphRunProperties endParagraphRunProperties2 = new A.EndParagraphRunProperties(){ Language = "en-US", Dirty = false };

            paragraph2.Append(run2);
            paragraph2.Append(endParagraphRunProperties2);

            textBody2.Append(bodyProperties2);
            textBody2.Append(listStyle2);
            textBody2.Append(paragraph2);

            shape2.Append(nonVisualShapeProperties2);
            shape2.Append(shapeProperties3);
            shape2.Append(textBody2);

            Shape shape3 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties3 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties7 = new NonVisualDrawingProperties(){ Id = (UInt32Value)4U, Name = "Date Placeholder 3" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties3 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks3 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties3.Append(shapeLocks3);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties7 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape3 = new PlaceholderShape(){ Type = PlaceholderValues.DateAndTime, Size = PlaceholderSizeValues.Half, Index = (UInt32Value)10U };

            applicationNonVisualDrawingProperties7.Append(placeholderShape3);

            nonVisualShapeProperties3.Append(nonVisualDrawingProperties7);
            nonVisualShapeProperties3.Append(nonVisualShapeDrawingProperties3);
            nonVisualShapeProperties3.Append(applicationNonVisualDrawingProperties7);
            ShapeProperties shapeProperties4 = new ShapeProperties();

            TextBody textBody3 = new TextBody();
            A.BodyProperties bodyProperties3 = new A.BodyProperties();
            A.ListStyle listStyle3 = new A.ListStyle();

            A.Paragraph paragraph3 = new A.Paragraph();

            A.Field field1 = new A.Field(){ Id = "{1D065981-0097-4BA9-A692-9FF652AFAF35}", Type = "datetimeFigureOut" };

            A.RunProperties runProperties3 = new A.RunProperties(){ Language = "cs-CZ" };
            runProperties3.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            A.Text text3 = new A.Text();
            text3.Text = "14.03.2023";

            field1.Append(runProperties3);
            field1.Append(text3);
            A.EndParagraphRunProperties endParagraphRunProperties3 = new A.EndParagraphRunProperties(){ Language = "cs-CZ" };

            paragraph3.Append(field1);
            paragraph3.Append(endParagraphRunProperties3);

            textBody3.Append(bodyProperties3);
            textBody3.Append(listStyle3);
            textBody3.Append(paragraph3);

            shape3.Append(nonVisualShapeProperties3);
            shape3.Append(shapeProperties4);
            shape3.Append(textBody3);

            Shape shape4 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties4 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties8 = new NonVisualDrawingProperties(){ Id = (UInt32Value)5U, Name = "Footer Placeholder 4" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties4 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks4 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties4.Append(shapeLocks4);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties8 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape4 = new PlaceholderShape(){ Type = PlaceholderValues.Footer, Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)11U };

            applicationNonVisualDrawingProperties8.Append(placeholderShape4);

            nonVisualShapeProperties4.Append(nonVisualDrawingProperties8);
            nonVisualShapeProperties4.Append(nonVisualShapeDrawingProperties4);
            nonVisualShapeProperties4.Append(applicationNonVisualDrawingProperties8);
            ShapeProperties shapeProperties5 = new ShapeProperties();

            TextBody textBody4 = new TextBody();
            A.BodyProperties bodyProperties4 = new A.BodyProperties();
            A.ListStyle listStyle4 = new A.ListStyle();

            A.Paragraph paragraph4 = new A.Paragraph();
            A.EndParagraphRunProperties endParagraphRunProperties4 = new A.EndParagraphRunProperties(){ Language = "cs-CZ" };

            paragraph4.Append(endParagraphRunProperties4);

            textBody4.Append(bodyProperties4);
            textBody4.Append(listStyle4);
            textBody4.Append(paragraph4);

            shape4.Append(nonVisualShapeProperties4);
            shape4.Append(shapeProperties5);
            shape4.Append(textBody4);

            Shape shape5 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties5 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties9 = new NonVisualDrawingProperties(){ Id = (UInt32Value)6U, Name = "Slide Number Placeholder 5" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties5 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks5 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties5.Append(shapeLocks5);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties9 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape5 = new PlaceholderShape(){ Type = PlaceholderValues.SlideNumber, Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)12U };

            applicationNonVisualDrawingProperties9.Append(placeholderShape5);

            nonVisualShapeProperties5.Append(nonVisualDrawingProperties9);
            nonVisualShapeProperties5.Append(nonVisualShapeDrawingProperties5);
            nonVisualShapeProperties5.Append(applicationNonVisualDrawingProperties9);
            ShapeProperties shapeProperties6 = new ShapeProperties();

            TextBody textBody5 = new TextBody();
            A.BodyProperties bodyProperties5 = new A.BodyProperties();
            A.ListStyle listStyle5 = new A.ListStyle();

            A.Paragraph paragraph5 = new A.Paragraph();

            A.Field field2 = new A.Field(){ Id = "{5D7F2B33-3E35-4F97-A448-9FEB151DF743}", Type = "slidenum" };

            A.RunProperties runProperties4 = new A.RunProperties(){ Language = "cs-CZ" };
            runProperties4.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            A.Text text4 = new A.Text();
            text4.Text = "‹#›";

            field2.Append(runProperties4);
            field2.Append(text4);
            A.EndParagraphRunProperties endParagraphRunProperties5 = new A.EndParagraphRunProperties(){ Language = "cs-CZ" };

            paragraph5.Append(field2);
            paragraph5.Append(endParagraphRunProperties5);

            textBody5.Append(bodyProperties5);
            textBody5.Append(listStyle5);
            textBody5.Append(paragraph5);

            shape5.Append(nonVisualShapeProperties5);
            shape5.Append(shapeProperties6);
            shape5.Append(textBody5);

            shapeTree2.Append(nonVisualGroupShapeProperties2);
            shapeTree2.Append(groupShapeProperties2);
            shapeTree2.Append(shape1);
            shapeTree2.Append(shape2);
            shapeTree2.Append(shape3);
            shapeTree2.Append(shape4);
            shapeTree2.Append(shape5);

            CommonSlideDataExtensionList commonSlideDataExtensionList2 = new CommonSlideDataExtensionList();

            CommonSlideDataExtension commonSlideDataExtension2 = new CommonSlideDataExtension(){ Uri = "{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}" };

            P14.CreationId creationId2 = new P14.CreationId(){ Val = (UInt32Value)3197096649U };
            creationId2.AddNamespaceDeclaration("p14", "http://schemas.microsoft.com/office/powerpoint/2010/main");

            commonSlideDataExtension2.Append(creationId2);

            commonSlideDataExtensionList2.Append(commonSlideDataExtension2);

            commonSlideData2.Append(shapeTree2);
            commonSlideData2.Append(commonSlideDataExtensionList2);

            ColorMapOverride colorMapOverride2 = new ColorMapOverride();
            A.MasterColorMapping masterColorMapping2 = new A.MasterColorMapping();

            colorMapOverride2.Append(masterColorMapping2);

            slideLayout1.Append(commonSlideData2);
            slideLayout1.Append(colorMapOverride2);

            slideLayoutPart1.SlideLayout = slideLayout1;
        }

        // Generates content of slideMasterPart1.
        private void GenerateSlideMasterPart1Content(SlideMasterPart slideMasterPart1)
        {
            SlideMaster slideMaster1 = new SlideMaster();
            slideMaster1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            slideMaster1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            slideMaster1.AddNamespaceDeclaration("p", "http://schemas.openxmlformats.org/presentationml/2006/main");

            CommonSlideData commonSlideData3 = new CommonSlideData();

            Background background1 = new Background();

            BackgroundStyleReference backgroundStyleReference1 = new BackgroundStyleReference(){ Index = (UInt32Value)1001U };
            A.SchemeColor schemeColor10 = new A.SchemeColor(){ Val = A.SchemeColorValues.Background1 };

            backgroundStyleReference1.Append(schemeColor10);

            background1.Append(backgroundStyleReference1);

            ShapeTree shapeTree3 = new ShapeTree();

            NonVisualGroupShapeProperties nonVisualGroupShapeProperties3 = new NonVisualGroupShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties10 = new NonVisualDrawingProperties(){ Id = (UInt32Value)1U, Name = "" };
            NonVisualGroupShapeDrawingProperties nonVisualGroupShapeDrawingProperties3 = new NonVisualGroupShapeDrawingProperties();
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties10 = new ApplicationNonVisualDrawingProperties();

            nonVisualGroupShapeProperties3.Append(nonVisualDrawingProperties10);
            nonVisualGroupShapeProperties3.Append(nonVisualGroupShapeDrawingProperties3);
            nonVisualGroupShapeProperties3.Append(applicationNonVisualDrawingProperties10);

            GroupShapeProperties groupShapeProperties3 = new GroupShapeProperties();

            A.TransformGroup transformGroup3 = new A.TransformGroup();
            A.Offset offset7 = new A.Offset(){ X = 0L, Y = 0L };
            A.Extents extents7 = new A.Extents(){ Cx = 0L, Cy = 0L };
            A.ChildOffset childOffset3 = new A.ChildOffset(){ X = 0L, Y = 0L };
            A.ChildExtents childExtents3 = new A.ChildExtents(){ Cx = 0L, Cy = 0L };

            transformGroup3.Append(offset7);
            transformGroup3.Append(extents7);
            transformGroup3.Append(childOffset3);
            transformGroup3.Append(childExtents3);

            groupShapeProperties3.Append(transformGroup3);

            Shape shape6 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties6 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties11 = new NonVisualDrawingProperties(){ Id = (UInt32Value)2U, Name = "Title Placeholder 1" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties6 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks6 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties6.Append(shapeLocks6);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties11 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape6 = new PlaceholderShape(){ Type = PlaceholderValues.Title };

            applicationNonVisualDrawingProperties11.Append(placeholderShape6);

            nonVisualShapeProperties6.Append(nonVisualDrawingProperties11);
            nonVisualShapeProperties6.Append(nonVisualShapeDrawingProperties6);
            nonVisualShapeProperties6.Append(applicationNonVisualDrawingProperties11);

            ShapeProperties shapeProperties7 = new ShapeProperties();

            A.Transform2D transform2D4 = new A.Transform2D();
            A.Offset offset8 = new A.Offset(){ X = 628650L, Y = 365126L };
            A.Extents extents8 = new A.Extents(){ Cx = 7886700L, Cy = 1325563L };

            transform2D4.Append(offset8);
            transform2D4.Append(extents8);

            A.PresetGeometry presetGeometry2 = new A.PresetGeometry(){ Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList2 = new A.AdjustValueList();

            presetGeometry2.Append(adjustValueList2);

            shapeProperties7.Append(transform2D4);
            shapeProperties7.Append(presetGeometry2);

            TextBody textBody6 = new TextBody();

            A.BodyProperties bodyProperties6 = new A.BodyProperties(){ Vertical = A.TextVerticalValues.Horizontal, LeftInset = 91440, TopInset = 45720, RightInset = 91440, BottomInset = 45720, RightToLeftColumns = false, Anchor = A.TextAnchoringTypeValues.Center };
            A.NormalAutoFit normalAutoFit1 = new A.NormalAutoFit();

            bodyProperties6.Append(normalAutoFit1);
            A.ListStyle listStyle6 = new A.ListStyle();

            A.Paragraph paragraph6 = new A.Paragraph();

            A.Run run3 = new A.Run();
            A.RunProperties runProperties5 = new A.RunProperties(){ Language = "cs-CZ" };
            A.Text text5 = new A.Text();
            text5.Text = "Kliknutím lze upravit styl.";

            run3.Append(runProperties5);
            run3.Append(text5);
            A.EndParagraphRunProperties endParagraphRunProperties6 = new A.EndParagraphRunProperties(){ Language = "en-US", Dirty = false };

            paragraph6.Append(run3);
            paragraph6.Append(endParagraphRunProperties6);

            textBody6.Append(bodyProperties6);
            textBody6.Append(listStyle6);
            textBody6.Append(paragraph6);

            shape6.Append(nonVisualShapeProperties6);
            shape6.Append(shapeProperties7);
            shape6.Append(textBody6);

            Shape shape7 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties7 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties12 = new NonVisualDrawingProperties(){ Id = (UInt32Value)3U, Name = "Text Placeholder 2" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties7 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks7 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties7.Append(shapeLocks7);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties12 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape7 = new PlaceholderShape(){ Type = PlaceholderValues.Body, Index = (UInt32Value)1U };

            applicationNonVisualDrawingProperties12.Append(placeholderShape7);

            nonVisualShapeProperties7.Append(nonVisualDrawingProperties12);
            nonVisualShapeProperties7.Append(nonVisualShapeDrawingProperties7);
            nonVisualShapeProperties7.Append(applicationNonVisualDrawingProperties12);

            ShapeProperties shapeProperties8 = new ShapeProperties();

            A.Transform2D transform2D5 = new A.Transform2D();
            A.Offset offset9 = new A.Offset(){ X = 628650L, Y = 1825625L };
            A.Extents extents9 = new A.Extents(){ Cx = 7886700L, Cy = 4351338L };

            transform2D5.Append(offset9);
            transform2D5.Append(extents9);

            A.PresetGeometry presetGeometry3 = new A.PresetGeometry(){ Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList3 = new A.AdjustValueList();

            presetGeometry3.Append(adjustValueList3);

            shapeProperties8.Append(transform2D5);
            shapeProperties8.Append(presetGeometry3);

            TextBody textBody7 = new TextBody();

            A.BodyProperties bodyProperties7 = new A.BodyProperties(){ Vertical = A.TextVerticalValues.Horizontal, LeftInset = 91440, TopInset = 45720, RightInset = 91440, BottomInset = 45720, RightToLeftColumns = false };
            A.NormalAutoFit normalAutoFit2 = new A.NormalAutoFit();

            bodyProperties7.Append(normalAutoFit2);
            A.ListStyle listStyle7 = new A.ListStyle();

            A.Paragraph paragraph7 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties1 = new A.ParagraphProperties(){ Level = 0 };

            A.Run run4 = new A.Run();
            A.RunProperties runProperties6 = new A.RunProperties(){ Language = "cs-CZ" };
            A.Text text6 = new A.Text();
            text6.Text = "Po kliknutí můžete upravovat styly textu v předloze.";

            run4.Append(runProperties6);
            run4.Append(text6);

            paragraph7.Append(paragraphProperties1);
            paragraph7.Append(run4);

            A.Paragraph paragraph8 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties2 = new A.ParagraphProperties(){ Level = 1 };

            A.Run run5 = new A.Run();
            A.RunProperties runProperties7 = new A.RunProperties(){ Language = "cs-CZ" };
            A.Text text7 = new A.Text();
            text7.Text = "Druhá úroveň";

            run5.Append(runProperties7);
            run5.Append(text7);

            paragraph8.Append(paragraphProperties2);
            paragraph8.Append(run5);

            A.Paragraph paragraph9 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties3 = new A.ParagraphProperties(){ Level = 2 };

            A.Run run6 = new A.Run();
            A.RunProperties runProperties8 = new A.RunProperties(){ Language = "cs-CZ" };
            A.Text text8 = new A.Text();
            text8.Text = "Třetí úroveň";

            run6.Append(runProperties8);
            run6.Append(text8);

            paragraph9.Append(paragraphProperties3);
            paragraph9.Append(run6);

            A.Paragraph paragraph10 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties4 = new A.ParagraphProperties(){ Level = 3 };

            A.Run run7 = new A.Run();
            A.RunProperties runProperties9 = new A.RunProperties(){ Language = "cs-CZ" };
            A.Text text9 = new A.Text();
            text9.Text = "Čtvrtá úroveň";

            run7.Append(runProperties9);
            run7.Append(text9);

            paragraph10.Append(paragraphProperties4);
            paragraph10.Append(run7);

            A.Paragraph paragraph11 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties5 = new A.ParagraphProperties(){ Level = 4 };

            A.Run run8 = new A.Run();
            A.RunProperties runProperties10 = new A.RunProperties(){ Language = "cs-CZ" };
            A.Text text10 = new A.Text();
            text10.Text = "Pátá úroveň";

            run8.Append(runProperties10);
            run8.Append(text10);
            A.EndParagraphRunProperties endParagraphRunProperties7 = new A.EndParagraphRunProperties(){ Language = "en-US", Dirty = false };

            paragraph11.Append(paragraphProperties5);
            paragraph11.Append(run8);
            paragraph11.Append(endParagraphRunProperties7);

            textBody7.Append(bodyProperties7);
            textBody7.Append(listStyle7);
            textBody7.Append(paragraph7);
            textBody7.Append(paragraph8);
            textBody7.Append(paragraph9);
            textBody7.Append(paragraph10);
            textBody7.Append(paragraph11);

            shape7.Append(nonVisualShapeProperties7);
            shape7.Append(shapeProperties8);
            shape7.Append(textBody7);

            Shape shape8 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties8 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties13 = new NonVisualDrawingProperties(){ Id = (UInt32Value)4U, Name = "Date Placeholder 3" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties8 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks8 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties8.Append(shapeLocks8);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties13 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape8 = new PlaceholderShape(){ Type = PlaceholderValues.DateAndTime, Size = PlaceholderSizeValues.Half, Index = (UInt32Value)2U };

            applicationNonVisualDrawingProperties13.Append(placeholderShape8);

            nonVisualShapeProperties8.Append(nonVisualDrawingProperties13);
            nonVisualShapeProperties8.Append(nonVisualShapeDrawingProperties8);
            nonVisualShapeProperties8.Append(applicationNonVisualDrawingProperties13);

            ShapeProperties shapeProperties9 = new ShapeProperties();

            A.Transform2D transform2D6 = new A.Transform2D();
            A.Offset offset10 = new A.Offset(){ X = 628650L, Y = 6356351L };
            A.Extents extents10 = new A.Extents(){ Cx = 2057400L, Cy = 365125L };

            transform2D6.Append(offset10);
            transform2D6.Append(extents10);

            A.PresetGeometry presetGeometry4 = new A.PresetGeometry(){ Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList4 = new A.AdjustValueList();

            presetGeometry4.Append(adjustValueList4);

            shapeProperties9.Append(transform2D6);
            shapeProperties9.Append(presetGeometry4);

            TextBody textBody8 = new TextBody();
            A.BodyProperties bodyProperties8 = new A.BodyProperties(){ Vertical = A.TextVerticalValues.Horizontal, LeftInset = 91440, TopInset = 45720, RightInset = 91440, BottomInset = 45720, RightToLeftColumns = false, Anchor = A.TextAnchoringTypeValues.Center };

            A.ListStyle listStyle8 = new A.ListStyle();

            A.Level1ParagraphProperties level1ParagraphProperties4 = new A.Level1ParagraphProperties(){ Alignment = A.TextAlignmentTypeValues.Left };

            A.DefaultRunProperties defaultRunProperties21 = new A.DefaultRunProperties(){ FontSize = 1200 };

            A.SolidFill solidFill10 = new A.SolidFill();

            A.SchemeColor schemeColor11 = new A.SchemeColor(){ Val = A.SchemeColorValues.Text1 };
            A.Tint tint1 = new A.Tint(){ Val = 75000 };

            schemeColor11.Append(tint1);

            solidFill10.Append(schemeColor11);

            defaultRunProperties21.Append(solidFill10);

            level1ParagraphProperties4.Append(defaultRunProperties21);

            listStyle8.Append(level1ParagraphProperties4);

            A.Paragraph paragraph12 = new A.Paragraph();

            A.Field field3 = new A.Field(){ Id = "{1D065981-0097-4BA9-A692-9FF652AFAF35}", Type = "datetimeFigureOut" };

            A.RunProperties runProperties11 = new A.RunProperties(){ Language = "cs-CZ" };
            runProperties11.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            A.Text text11 = new A.Text();
            text11.Text = "14.03.2023";

            field3.Append(runProperties11);
            field3.Append(text11);
            A.EndParagraphRunProperties endParagraphRunProperties8 = new A.EndParagraphRunProperties(){ Language = "cs-CZ" };

            paragraph12.Append(field3);
            paragraph12.Append(endParagraphRunProperties8);

            textBody8.Append(bodyProperties8);
            textBody8.Append(listStyle8);
            textBody8.Append(paragraph12);

            shape8.Append(nonVisualShapeProperties8);
            shape8.Append(shapeProperties9);
            shape8.Append(textBody8);

            Shape shape9 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties9 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties14 = new NonVisualDrawingProperties(){ Id = (UInt32Value)5U, Name = "Footer Placeholder 4" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties9 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks9 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties9.Append(shapeLocks9);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties14 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape9 = new PlaceholderShape(){ Type = PlaceholderValues.Footer, Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)3U };

            applicationNonVisualDrawingProperties14.Append(placeholderShape9);

            nonVisualShapeProperties9.Append(nonVisualDrawingProperties14);
            nonVisualShapeProperties9.Append(nonVisualShapeDrawingProperties9);
            nonVisualShapeProperties9.Append(applicationNonVisualDrawingProperties14);

            ShapeProperties shapeProperties10 = new ShapeProperties();

            A.Transform2D transform2D7 = new A.Transform2D();
            A.Offset offset11 = new A.Offset(){ X = 3028950L, Y = 6356351L };
            A.Extents extents11 = new A.Extents(){ Cx = 3086100L, Cy = 365125L };

            transform2D7.Append(offset11);
            transform2D7.Append(extents11);

            A.PresetGeometry presetGeometry5 = new A.PresetGeometry(){ Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList5 = new A.AdjustValueList();

            presetGeometry5.Append(adjustValueList5);

            shapeProperties10.Append(transform2D7);
            shapeProperties10.Append(presetGeometry5);

            TextBody textBody9 = new TextBody();
            A.BodyProperties bodyProperties9 = new A.BodyProperties(){ Vertical = A.TextVerticalValues.Horizontal, LeftInset = 91440, TopInset = 45720, RightInset = 91440, BottomInset = 45720, RightToLeftColumns = false, Anchor = A.TextAnchoringTypeValues.Center };

            A.ListStyle listStyle9 = new A.ListStyle();

            A.Level1ParagraphProperties level1ParagraphProperties5 = new A.Level1ParagraphProperties(){ Alignment = A.TextAlignmentTypeValues.Center };

            A.DefaultRunProperties defaultRunProperties22 = new A.DefaultRunProperties(){ FontSize = 1200 };

            A.SolidFill solidFill11 = new A.SolidFill();

            A.SchemeColor schemeColor12 = new A.SchemeColor(){ Val = A.SchemeColorValues.Text1 };
            A.Tint tint2 = new A.Tint(){ Val = 75000 };

            schemeColor12.Append(tint2);

            solidFill11.Append(schemeColor12);

            defaultRunProperties22.Append(solidFill11);

            level1ParagraphProperties5.Append(defaultRunProperties22);

            listStyle9.Append(level1ParagraphProperties5);

            A.Paragraph paragraph13 = new A.Paragraph();
            A.EndParagraphRunProperties endParagraphRunProperties9 = new A.EndParagraphRunProperties(){ Language = "cs-CZ" };

            paragraph13.Append(endParagraphRunProperties9);

            textBody9.Append(bodyProperties9);
            textBody9.Append(listStyle9);
            textBody9.Append(paragraph13);

            shape9.Append(nonVisualShapeProperties9);
            shape9.Append(shapeProperties10);
            shape9.Append(textBody9);

            Shape shape10 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties10 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties15 = new NonVisualDrawingProperties(){ Id = (UInt32Value)6U, Name = "Slide Number Placeholder 5" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties10 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks10 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties10.Append(shapeLocks10);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties15 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape10 = new PlaceholderShape(){ Type = PlaceholderValues.SlideNumber, Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)4U };

            applicationNonVisualDrawingProperties15.Append(placeholderShape10);

            nonVisualShapeProperties10.Append(nonVisualDrawingProperties15);
            nonVisualShapeProperties10.Append(nonVisualShapeDrawingProperties10);
            nonVisualShapeProperties10.Append(applicationNonVisualDrawingProperties15);

            ShapeProperties shapeProperties11 = new ShapeProperties();

            A.Transform2D transform2D8 = new A.Transform2D();
            A.Offset offset12 = new A.Offset(){ X = 6457950L, Y = 6356351L };
            A.Extents extents12 = new A.Extents(){ Cx = 2057400L, Cy = 365125L };

            transform2D8.Append(offset12);
            transform2D8.Append(extents12);

            A.PresetGeometry presetGeometry6 = new A.PresetGeometry(){ Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList6 = new A.AdjustValueList();

            presetGeometry6.Append(adjustValueList6);

            shapeProperties11.Append(transform2D8);
            shapeProperties11.Append(presetGeometry6);

            TextBody textBody10 = new TextBody();
            A.BodyProperties bodyProperties10 = new A.BodyProperties(){ Vertical = A.TextVerticalValues.Horizontal, LeftInset = 91440, TopInset = 45720, RightInset = 91440, BottomInset = 45720, RightToLeftColumns = false, Anchor = A.TextAnchoringTypeValues.Center };

            A.ListStyle listStyle10 = new A.ListStyle();

            A.Level1ParagraphProperties level1ParagraphProperties6 = new A.Level1ParagraphProperties(){ Alignment = A.TextAlignmentTypeValues.Right };

            A.DefaultRunProperties defaultRunProperties23 = new A.DefaultRunProperties(){ FontSize = 1200 };

            A.SolidFill solidFill12 = new A.SolidFill();

            A.SchemeColor schemeColor13 = new A.SchemeColor(){ Val = A.SchemeColorValues.Text1 };
            A.Tint tint3 = new A.Tint(){ Val = 75000 };

            schemeColor13.Append(tint3);

            solidFill12.Append(schemeColor13);

            defaultRunProperties23.Append(solidFill12);

            level1ParagraphProperties6.Append(defaultRunProperties23);

            listStyle10.Append(level1ParagraphProperties6);

            A.Paragraph paragraph14 = new A.Paragraph();

            A.Field field4 = new A.Field(){ Id = "{5D7F2B33-3E35-4F97-A448-9FEB151DF743}", Type = "slidenum" };

            A.RunProperties runProperties12 = new A.RunProperties(){ Language = "cs-CZ" };
            runProperties12.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            A.Text text12 = new A.Text();
            text12.Text = "‹#›";

            field4.Append(runProperties12);
            field4.Append(text12);
            A.EndParagraphRunProperties endParagraphRunProperties10 = new A.EndParagraphRunProperties(){ Language = "cs-CZ" };

            paragraph14.Append(field4);
            paragraph14.Append(endParagraphRunProperties10);

            textBody10.Append(bodyProperties10);
            textBody10.Append(listStyle10);
            textBody10.Append(paragraph14);

            shape10.Append(nonVisualShapeProperties10);
            shape10.Append(shapeProperties11);
            shape10.Append(textBody10);

            shapeTree3.Append(nonVisualGroupShapeProperties3);
            shapeTree3.Append(groupShapeProperties3);
            shapeTree3.Append(shape6);
            shapeTree3.Append(shape7);
            shapeTree3.Append(shape8);
            shapeTree3.Append(shape9);
            shapeTree3.Append(shape10);

            CommonSlideDataExtensionList commonSlideDataExtensionList3 = new CommonSlideDataExtensionList();

            CommonSlideDataExtension commonSlideDataExtension3 = new CommonSlideDataExtension(){ Uri = "{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}" };

            P14.CreationId creationId3 = new P14.CreationId(){ Val = (UInt32Value)1840157969U };
            creationId3.AddNamespaceDeclaration("p14", "http://schemas.microsoft.com/office/powerpoint/2010/main");

            commonSlideDataExtension3.Append(creationId3);

            commonSlideDataExtensionList3.Append(commonSlideDataExtension3);

            commonSlideData3.Append(background1);
            commonSlideData3.Append(shapeTree3);
            commonSlideData3.Append(commonSlideDataExtensionList3);
            ColorMap colorMap1 = new ColorMap(){ Background1 = A.ColorSchemeIndexValues.Light1, Text1 = A.ColorSchemeIndexValues.Dark1, Background2 = A.ColorSchemeIndexValues.Light2, Text2 = A.ColorSchemeIndexValues.Dark2, Accent1 = A.ColorSchemeIndexValues.Accent1, Accent2 = A.ColorSchemeIndexValues.Accent2, Accent3 = A.ColorSchemeIndexValues.Accent3, Accent4 = A.ColorSchemeIndexValues.Accent4, Accent5 = A.ColorSchemeIndexValues.Accent5, Accent6 = A.ColorSchemeIndexValues.Accent6, Hyperlink = A.ColorSchemeIndexValues.Hyperlink, FollowedHyperlink = A.ColorSchemeIndexValues.FollowedHyperlink };

            SlideLayoutIdList slideLayoutIdList1 = new SlideLayoutIdList();
            SlideLayoutId slideLayoutId1 = new SlideLayoutId(){ Id = (UInt32Value)2147483661U, RelationshipId = "rId1" };
            SlideLayoutId slideLayoutId2 = new SlideLayoutId(){ Id = (UInt32Value)2147483662U, RelationshipId = "rId2" };
            SlideLayoutId slideLayoutId3 = new SlideLayoutId(){ Id = (UInt32Value)2147483663U, RelationshipId = "rId3" };
            SlideLayoutId slideLayoutId4 = new SlideLayoutId(){ Id = (UInt32Value)2147483664U, RelationshipId = "rId4" };
            SlideLayoutId slideLayoutId5 = new SlideLayoutId(){ Id = (UInt32Value)2147483665U, RelationshipId = "rId5" };
            SlideLayoutId slideLayoutId6 = new SlideLayoutId(){ Id = (UInt32Value)2147483666U, RelationshipId = "rId6" };
            SlideLayoutId slideLayoutId7 = new SlideLayoutId(){ Id = (UInt32Value)2147483667U, RelationshipId = "rId7" };
            SlideLayoutId slideLayoutId8 = new SlideLayoutId(){ Id = (UInt32Value)2147483668U, RelationshipId = "rId8" };
            SlideLayoutId slideLayoutId9 = new SlideLayoutId(){ Id = (UInt32Value)2147483669U, RelationshipId = "rId9" };
            SlideLayoutId slideLayoutId10 = new SlideLayoutId(){ Id = (UInt32Value)2147483670U, RelationshipId = "rId10" };
            SlideLayoutId slideLayoutId11 = new SlideLayoutId(){ Id = (UInt32Value)2147483671U, RelationshipId = "rId11" };

            slideLayoutIdList1.Append(slideLayoutId1);
            slideLayoutIdList1.Append(slideLayoutId2);
            slideLayoutIdList1.Append(slideLayoutId3);
            slideLayoutIdList1.Append(slideLayoutId4);
            slideLayoutIdList1.Append(slideLayoutId5);
            slideLayoutIdList1.Append(slideLayoutId6);
            slideLayoutIdList1.Append(slideLayoutId7);
            slideLayoutIdList1.Append(slideLayoutId8);
            slideLayoutIdList1.Append(slideLayoutId9);
            slideLayoutIdList1.Append(slideLayoutId10);
            slideLayoutIdList1.Append(slideLayoutId11);

            TextStyles textStyles1 = new TextStyles();

            TitleStyle titleStyle1 = new TitleStyle();

            A.Level1ParagraphProperties level1ParagraphProperties7 = new A.Level1ParagraphProperties(){ Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.LineSpacing lineSpacing1 = new A.LineSpacing();
            A.SpacingPercent spacingPercent1 = new A.SpacingPercent(){ Val = 90000 };

            lineSpacing1.Append(spacingPercent1);

            A.SpaceBefore spaceBefore1 = new A.SpaceBefore();
            A.SpacingPercent spacingPercent2 = new A.SpacingPercent(){ Val = 0 };

            spaceBefore1.Append(spacingPercent2);
            A.NoBullet noBullet10 = new A.NoBullet();

            A.DefaultRunProperties defaultRunProperties24 = new A.DefaultRunProperties(){ FontSize = 4400, Kerning = 1200 };

            A.SolidFill solidFill13 = new A.SolidFill();
            A.SchemeColor schemeColor14 = new A.SchemeColor(){ Val = A.SchemeColorValues.Text1 };

            solidFill13.Append(schemeColor14);
            A.LatinFont latinFont10 = new A.LatinFont(){ Typeface = "+mj-lt" };
            A.EastAsianFont eastAsianFont10 = new A.EastAsianFont(){ Typeface = "+mj-ea" };
            A.ComplexScriptFont complexScriptFont10 = new A.ComplexScriptFont(){ Typeface = "+mj-cs" };

            defaultRunProperties24.Append(solidFill13);
            defaultRunProperties24.Append(latinFont10);
            defaultRunProperties24.Append(eastAsianFont10);
            defaultRunProperties24.Append(complexScriptFont10);

            level1ParagraphProperties7.Append(lineSpacing1);
            level1ParagraphProperties7.Append(spaceBefore1);
            level1ParagraphProperties7.Append(noBullet10);
            level1ParagraphProperties7.Append(defaultRunProperties24);

            titleStyle1.Append(level1ParagraphProperties7);

            BodyStyle bodyStyle1 = new BodyStyle();

            A.Level1ParagraphProperties level1ParagraphProperties8 = new A.Level1ParagraphProperties(){ LeftMargin = 228600, Indent = -228600, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.LineSpacing lineSpacing2 = new A.LineSpacing();
            A.SpacingPercent spacingPercent3 = new A.SpacingPercent(){ Val = 90000 };

            lineSpacing2.Append(spacingPercent3);

            A.SpaceBefore spaceBefore2 = new A.SpaceBefore();
            A.SpacingPoints spacingPoints1 = new A.SpacingPoints(){ Val = 1000 };

            spaceBefore2.Append(spacingPoints1);
            A.BulletFont bulletFont1 = new A.BulletFont(){ Typeface = "Arial", Panose = "020B0604020202020204", PitchFamily = 34, CharacterSet = 0 };
            A.CharacterBullet characterBullet1 = new A.CharacterBullet(){ Char = "•" };

            A.DefaultRunProperties defaultRunProperties25 = new A.DefaultRunProperties(){ FontSize = 2800, Kerning = 1200 };

            A.SolidFill solidFill14 = new A.SolidFill();
            A.SchemeColor schemeColor15 = new A.SchemeColor(){ Val = A.SchemeColorValues.Text1 };

            solidFill14.Append(schemeColor15);
            A.LatinFont latinFont11 = new A.LatinFont(){ Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont11 = new A.EastAsianFont(){ Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont11 = new A.ComplexScriptFont(){ Typeface = "+mn-cs" };

            defaultRunProperties25.Append(solidFill14);
            defaultRunProperties25.Append(latinFont11);
            defaultRunProperties25.Append(eastAsianFont11);
            defaultRunProperties25.Append(complexScriptFont11);

            level1ParagraphProperties8.Append(lineSpacing2);
            level1ParagraphProperties8.Append(spaceBefore2);
            level1ParagraphProperties8.Append(bulletFont1);
            level1ParagraphProperties8.Append(characterBullet1);
            level1ParagraphProperties8.Append(defaultRunProperties25);

            A.Level2ParagraphProperties level2ParagraphProperties3 = new A.Level2ParagraphProperties(){ LeftMargin = 685800, Indent = -228600, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.LineSpacing lineSpacing3 = new A.LineSpacing();
            A.SpacingPercent spacingPercent4 = new A.SpacingPercent(){ Val = 90000 };

            lineSpacing3.Append(spacingPercent4);

            A.SpaceBefore spaceBefore3 = new A.SpaceBefore();
            A.SpacingPoints spacingPoints2 = new A.SpacingPoints(){ Val = 500 };

            spaceBefore3.Append(spacingPoints2);
            A.BulletFont bulletFont2 = new A.BulletFont(){ Typeface = "Arial", Panose = "020B0604020202020204", PitchFamily = 34, CharacterSet = 0 };
            A.CharacterBullet characterBullet2 = new A.CharacterBullet(){ Char = "•" };

            A.DefaultRunProperties defaultRunProperties26 = new A.DefaultRunProperties(){ FontSize = 2400, Kerning = 1200 };

            A.SolidFill solidFill15 = new A.SolidFill();
            A.SchemeColor schemeColor16 = new A.SchemeColor(){ Val = A.SchemeColorValues.Text1 };

            solidFill15.Append(schemeColor16);
            A.LatinFont latinFont12 = new A.LatinFont(){ Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont12 = new A.EastAsianFont(){ Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont12 = new A.ComplexScriptFont(){ Typeface = "+mn-cs" };

            defaultRunProperties26.Append(solidFill15);
            defaultRunProperties26.Append(latinFont12);
            defaultRunProperties26.Append(eastAsianFont12);
            defaultRunProperties26.Append(complexScriptFont12);

            level2ParagraphProperties3.Append(lineSpacing3);
            level2ParagraphProperties3.Append(spaceBefore3);
            level2ParagraphProperties3.Append(bulletFont2);
            level2ParagraphProperties3.Append(characterBullet2);
            level2ParagraphProperties3.Append(defaultRunProperties26);

            A.Level3ParagraphProperties level3ParagraphProperties3 = new A.Level3ParagraphProperties(){ LeftMargin = 1143000, Indent = -228600, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.LineSpacing lineSpacing4 = new A.LineSpacing();
            A.SpacingPercent spacingPercent5 = new A.SpacingPercent(){ Val = 90000 };

            lineSpacing4.Append(spacingPercent5);

            A.SpaceBefore spaceBefore4 = new A.SpaceBefore();
            A.SpacingPoints spacingPoints3 = new A.SpacingPoints(){ Val = 500 };

            spaceBefore4.Append(spacingPoints3);
            A.BulletFont bulletFont3 = new A.BulletFont(){ Typeface = "Arial", Panose = "020B0604020202020204", PitchFamily = 34, CharacterSet = 0 };
            A.CharacterBullet characterBullet3 = new A.CharacterBullet(){ Char = "•" };

            A.DefaultRunProperties defaultRunProperties27 = new A.DefaultRunProperties(){ FontSize = 2000, Kerning = 1200 };

            A.SolidFill solidFill16 = new A.SolidFill();
            A.SchemeColor schemeColor17 = new A.SchemeColor(){ Val = A.SchemeColorValues.Text1 };

            solidFill16.Append(schemeColor17);
            A.LatinFont latinFont13 = new A.LatinFont(){ Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont13 = new A.EastAsianFont(){ Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont13 = new A.ComplexScriptFont(){ Typeface = "+mn-cs" };

            defaultRunProperties27.Append(solidFill16);
            defaultRunProperties27.Append(latinFont13);
            defaultRunProperties27.Append(eastAsianFont13);
            defaultRunProperties27.Append(complexScriptFont13);

            level3ParagraphProperties3.Append(lineSpacing4);
            level3ParagraphProperties3.Append(spaceBefore4);
            level3ParagraphProperties3.Append(bulletFont3);
            level3ParagraphProperties3.Append(characterBullet3);
            level3ParagraphProperties3.Append(defaultRunProperties27);

            A.Level4ParagraphProperties level4ParagraphProperties3 = new A.Level4ParagraphProperties(){ LeftMargin = 1600200, Indent = -228600, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.LineSpacing lineSpacing5 = new A.LineSpacing();
            A.SpacingPercent spacingPercent6 = new A.SpacingPercent(){ Val = 90000 };

            lineSpacing5.Append(spacingPercent6);

            A.SpaceBefore spaceBefore5 = new A.SpaceBefore();
            A.SpacingPoints spacingPoints4 = new A.SpacingPoints(){ Val = 500 };

            spaceBefore5.Append(spacingPoints4);
            A.BulletFont bulletFont4 = new A.BulletFont(){ Typeface = "Arial", Panose = "020B0604020202020204", PitchFamily = 34, CharacterSet = 0 };
            A.CharacterBullet characterBullet4 = new A.CharacterBullet(){ Char = "•" };

            A.DefaultRunProperties defaultRunProperties28 = new A.DefaultRunProperties(){ FontSize = 1800, Kerning = 1200 };

            A.SolidFill solidFill17 = new A.SolidFill();
            A.SchemeColor schemeColor18 = new A.SchemeColor(){ Val = A.SchemeColorValues.Text1 };

            solidFill17.Append(schemeColor18);
            A.LatinFont latinFont14 = new A.LatinFont(){ Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont14 = new A.EastAsianFont(){ Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont14 = new A.ComplexScriptFont(){ Typeface = "+mn-cs" };

            defaultRunProperties28.Append(solidFill17);
            defaultRunProperties28.Append(latinFont14);
            defaultRunProperties28.Append(eastAsianFont14);
            defaultRunProperties28.Append(complexScriptFont14);

            level4ParagraphProperties3.Append(lineSpacing5);
            level4ParagraphProperties3.Append(spaceBefore5);
            level4ParagraphProperties3.Append(bulletFont4);
            level4ParagraphProperties3.Append(characterBullet4);
            level4ParagraphProperties3.Append(defaultRunProperties28);

            A.Level5ParagraphProperties level5ParagraphProperties3 = new A.Level5ParagraphProperties(){ LeftMargin = 2057400, Indent = -228600, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.LineSpacing lineSpacing6 = new A.LineSpacing();
            A.SpacingPercent spacingPercent7 = new A.SpacingPercent(){ Val = 90000 };

            lineSpacing6.Append(spacingPercent7);

            A.SpaceBefore spaceBefore6 = new A.SpaceBefore();
            A.SpacingPoints spacingPoints5 = new A.SpacingPoints(){ Val = 500 };

            spaceBefore6.Append(spacingPoints5);
            A.BulletFont bulletFont5 = new A.BulletFont(){ Typeface = "Arial", Panose = "020B0604020202020204", PitchFamily = 34, CharacterSet = 0 };
            A.CharacterBullet characterBullet5 = new A.CharacterBullet(){ Char = "•" };

            A.DefaultRunProperties defaultRunProperties29 = new A.DefaultRunProperties(){ FontSize = 1800, Kerning = 1200 };

            A.SolidFill solidFill18 = new A.SolidFill();
            A.SchemeColor schemeColor19 = new A.SchemeColor(){ Val = A.SchemeColorValues.Text1 };

            solidFill18.Append(schemeColor19);
            A.LatinFont latinFont15 = new A.LatinFont(){ Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont15 = new A.EastAsianFont(){ Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont15 = new A.ComplexScriptFont(){ Typeface = "+mn-cs" };

            defaultRunProperties29.Append(solidFill18);
            defaultRunProperties29.Append(latinFont15);
            defaultRunProperties29.Append(eastAsianFont15);
            defaultRunProperties29.Append(complexScriptFont15);

            level5ParagraphProperties3.Append(lineSpacing6);
            level5ParagraphProperties3.Append(spaceBefore6);
            level5ParagraphProperties3.Append(bulletFont5);
            level5ParagraphProperties3.Append(characterBullet5);
            level5ParagraphProperties3.Append(defaultRunProperties29);

            A.Level6ParagraphProperties level6ParagraphProperties3 = new A.Level6ParagraphProperties(){ LeftMargin = 2514600, Indent = -228600, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.LineSpacing lineSpacing7 = new A.LineSpacing();
            A.SpacingPercent spacingPercent8 = new A.SpacingPercent(){ Val = 90000 };

            lineSpacing7.Append(spacingPercent8);

            A.SpaceBefore spaceBefore7 = new A.SpaceBefore();
            A.SpacingPoints spacingPoints6 = new A.SpacingPoints(){ Val = 500 };

            spaceBefore7.Append(spacingPoints6);
            A.BulletFont bulletFont6 = new A.BulletFont(){ Typeface = "Arial", Panose = "020B0604020202020204", PitchFamily = 34, CharacterSet = 0 };
            A.CharacterBullet characterBullet6 = new A.CharacterBullet(){ Char = "•" };

            A.DefaultRunProperties defaultRunProperties30 = new A.DefaultRunProperties(){ FontSize = 1800, Kerning = 1200 };

            A.SolidFill solidFill19 = new A.SolidFill();
            A.SchemeColor schemeColor20 = new A.SchemeColor(){ Val = A.SchemeColorValues.Text1 };

            solidFill19.Append(schemeColor20);
            A.LatinFont latinFont16 = new A.LatinFont(){ Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont16 = new A.EastAsianFont(){ Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont16 = new A.ComplexScriptFont(){ Typeface = "+mn-cs" };

            defaultRunProperties30.Append(solidFill19);
            defaultRunProperties30.Append(latinFont16);
            defaultRunProperties30.Append(eastAsianFont16);
            defaultRunProperties30.Append(complexScriptFont16);

            level6ParagraphProperties3.Append(lineSpacing7);
            level6ParagraphProperties3.Append(spaceBefore7);
            level6ParagraphProperties3.Append(bulletFont6);
            level6ParagraphProperties3.Append(characterBullet6);
            level6ParagraphProperties3.Append(defaultRunProperties30);

            A.Level7ParagraphProperties level7ParagraphProperties3 = new A.Level7ParagraphProperties(){ LeftMargin = 2971800, Indent = -228600, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.LineSpacing lineSpacing8 = new A.LineSpacing();
            A.SpacingPercent spacingPercent9 = new A.SpacingPercent(){ Val = 90000 };

            lineSpacing8.Append(spacingPercent9);

            A.SpaceBefore spaceBefore8 = new A.SpaceBefore();
            A.SpacingPoints spacingPoints7 = new A.SpacingPoints(){ Val = 500 };

            spaceBefore8.Append(spacingPoints7);
            A.BulletFont bulletFont7 = new A.BulletFont(){ Typeface = "Arial", Panose = "020B0604020202020204", PitchFamily = 34, CharacterSet = 0 };
            A.CharacterBullet characterBullet7 = new A.CharacterBullet(){ Char = "•" };

            A.DefaultRunProperties defaultRunProperties31 = new A.DefaultRunProperties(){ FontSize = 1800, Kerning = 1200 };

            A.SolidFill solidFill20 = new A.SolidFill();
            A.SchemeColor schemeColor21 = new A.SchemeColor(){ Val = A.SchemeColorValues.Text1 };

            solidFill20.Append(schemeColor21);
            A.LatinFont latinFont17 = new A.LatinFont(){ Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont17 = new A.EastAsianFont(){ Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont17 = new A.ComplexScriptFont(){ Typeface = "+mn-cs" };

            defaultRunProperties31.Append(solidFill20);
            defaultRunProperties31.Append(latinFont17);
            defaultRunProperties31.Append(eastAsianFont17);
            defaultRunProperties31.Append(complexScriptFont17);

            level7ParagraphProperties3.Append(lineSpacing8);
            level7ParagraphProperties3.Append(spaceBefore8);
            level7ParagraphProperties3.Append(bulletFont7);
            level7ParagraphProperties3.Append(characterBullet7);
            level7ParagraphProperties3.Append(defaultRunProperties31);

            A.Level8ParagraphProperties level8ParagraphProperties3 = new A.Level8ParagraphProperties(){ LeftMargin = 3429000, Indent = -228600, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.LineSpacing lineSpacing9 = new A.LineSpacing();
            A.SpacingPercent spacingPercent10 = new A.SpacingPercent(){ Val = 90000 };

            lineSpacing9.Append(spacingPercent10);

            A.SpaceBefore spaceBefore9 = new A.SpaceBefore();
            A.SpacingPoints spacingPoints8 = new A.SpacingPoints(){ Val = 500 };

            spaceBefore9.Append(spacingPoints8);
            A.BulletFont bulletFont8 = new A.BulletFont(){ Typeface = "Arial", Panose = "020B0604020202020204", PitchFamily = 34, CharacterSet = 0 };
            A.CharacterBullet characterBullet8 = new A.CharacterBullet(){ Char = "•" };

            A.DefaultRunProperties defaultRunProperties32 = new A.DefaultRunProperties(){ FontSize = 1800, Kerning = 1200 };

            A.SolidFill solidFill21 = new A.SolidFill();
            A.SchemeColor schemeColor22 = new A.SchemeColor(){ Val = A.SchemeColorValues.Text1 };

            solidFill21.Append(schemeColor22);
            A.LatinFont latinFont18 = new A.LatinFont(){ Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont18 = new A.EastAsianFont(){ Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont18 = new A.ComplexScriptFont(){ Typeface = "+mn-cs" };

            defaultRunProperties32.Append(solidFill21);
            defaultRunProperties32.Append(latinFont18);
            defaultRunProperties32.Append(eastAsianFont18);
            defaultRunProperties32.Append(complexScriptFont18);

            level8ParagraphProperties3.Append(lineSpacing9);
            level8ParagraphProperties3.Append(spaceBefore9);
            level8ParagraphProperties3.Append(bulletFont8);
            level8ParagraphProperties3.Append(characterBullet8);
            level8ParagraphProperties3.Append(defaultRunProperties32);

            A.Level9ParagraphProperties level9ParagraphProperties3 = new A.Level9ParagraphProperties(){ LeftMargin = 3886200, Indent = -228600, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.LineSpacing lineSpacing10 = new A.LineSpacing();
            A.SpacingPercent spacingPercent11 = new A.SpacingPercent(){ Val = 90000 };

            lineSpacing10.Append(spacingPercent11);

            A.SpaceBefore spaceBefore10 = new A.SpaceBefore();
            A.SpacingPoints spacingPoints9 = new A.SpacingPoints(){ Val = 500 };

            spaceBefore10.Append(spacingPoints9);
            A.BulletFont bulletFont9 = new A.BulletFont(){ Typeface = "Arial", Panose = "020B0604020202020204", PitchFamily = 34, CharacterSet = 0 };
            A.CharacterBullet characterBullet9 = new A.CharacterBullet(){ Char = "•" };

            A.DefaultRunProperties defaultRunProperties33 = new A.DefaultRunProperties(){ FontSize = 1800, Kerning = 1200 };

            A.SolidFill solidFill22 = new A.SolidFill();
            A.SchemeColor schemeColor23 = new A.SchemeColor(){ Val = A.SchemeColorValues.Text1 };

            solidFill22.Append(schemeColor23);
            A.LatinFont latinFont19 = new A.LatinFont(){ Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont19 = new A.EastAsianFont(){ Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont19 = new A.ComplexScriptFont(){ Typeface = "+mn-cs" };

            defaultRunProperties33.Append(solidFill22);
            defaultRunProperties33.Append(latinFont19);
            defaultRunProperties33.Append(eastAsianFont19);
            defaultRunProperties33.Append(complexScriptFont19);

            level9ParagraphProperties3.Append(lineSpacing10);
            level9ParagraphProperties3.Append(spaceBefore10);
            level9ParagraphProperties3.Append(bulletFont9);
            level9ParagraphProperties3.Append(characterBullet9);
            level9ParagraphProperties3.Append(defaultRunProperties33);

            bodyStyle1.Append(level1ParagraphProperties8);
            bodyStyle1.Append(level2ParagraphProperties3);
            bodyStyle1.Append(level3ParagraphProperties3);
            bodyStyle1.Append(level4ParagraphProperties3);
            bodyStyle1.Append(level5ParagraphProperties3);
            bodyStyle1.Append(level6ParagraphProperties3);
            bodyStyle1.Append(level7ParagraphProperties3);
            bodyStyle1.Append(level8ParagraphProperties3);
            bodyStyle1.Append(level9ParagraphProperties3);

            OtherStyle otherStyle1 = new OtherStyle();

            A.DefaultParagraphProperties defaultParagraphProperties2 = new A.DefaultParagraphProperties();
            A.DefaultRunProperties defaultRunProperties34 = new A.DefaultRunProperties(){ Language = "en-US" };

            defaultParagraphProperties2.Append(defaultRunProperties34);

            A.Level1ParagraphProperties level1ParagraphProperties9 = new A.Level1ParagraphProperties(){ LeftMargin = 0, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.DefaultRunProperties defaultRunProperties35 = new A.DefaultRunProperties(){ FontSize = 1800, Kerning = 1200 };

            A.SolidFill solidFill23 = new A.SolidFill();
            A.SchemeColor schemeColor24 = new A.SchemeColor(){ Val = A.SchemeColorValues.Text1 };

            solidFill23.Append(schemeColor24);
            A.LatinFont latinFont20 = new A.LatinFont(){ Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont20 = new A.EastAsianFont(){ Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont20 = new A.ComplexScriptFont(){ Typeface = "+mn-cs" };

            defaultRunProperties35.Append(solidFill23);
            defaultRunProperties35.Append(latinFont20);
            defaultRunProperties35.Append(eastAsianFont20);
            defaultRunProperties35.Append(complexScriptFont20);

            level1ParagraphProperties9.Append(defaultRunProperties35);

            A.Level2ParagraphProperties level2ParagraphProperties4 = new A.Level2ParagraphProperties(){ LeftMargin = 457200, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.DefaultRunProperties defaultRunProperties36 = new A.DefaultRunProperties(){ FontSize = 1800, Kerning = 1200 };

            A.SolidFill solidFill24 = new A.SolidFill();
            A.SchemeColor schemeColor25 = new A.SchemeColor(){ Val = A.SchemeColorValues.Text1 };

            solidFill24.Append(schemeColor25);
            A.LatinFont latinFont21 = new A.LatinFont(){ Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont21 = new A.EastAsianFont(){ Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont21 = new A.ComplexScriptFont(){ Typeface = "+mn-cs" };

            defaultRunProperties36.Append(solidFill24);
            defaultRunProperties36.Append(latinFont21);
            defaultRunProperties36.Append(eastAsianFont21);
            defaultRunProperties36.Append(complexScriptFont21);

            level2ParagraphProperties4.Append(defaultRunProperties36);

            A.Level3ParagraphProperties level3ParagraphProperties4 = new A.Level3ParagraphProperties(){ LeftMargin = 914400, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.DefaultRunProperties defaultRunProperties37 = new A.DefaultRunProperties(){ FontSize = 1800, Kerning = 1200 };

            A.SolidFill solidFill25 = new A.SolidFill();
            A.SchemeColor schemeColor26 = new A.SchemeColor(){ Val = A.SchemeColorValues.Text1 };

            solidFill25.Append(schemeColor26);
            A.LatinFont latinFont22 = new A.LatinFont(){ Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont22 = new A.EastAsianFont(){ Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont22 = new A.ComplexScriptFont(){ Typeface = "+mn-cs" };

            defaultRunProperties37.Append(solidFill25);
            defaultRunProperties37.Append(latinFont22);
            defaultRunProperties37.Append(eastAsianFont22);
            defaultRunProperties37.Append(complexScriptFont22);

            level3ParagraphProperties4.Append(defaultRunProperties37);

            A.Level4ParagraphProperties level4ParagraphProperties4 = new A.Level4ParagraphProperties(){ LeftMargin = 1371600, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.DefaultRunProperties defaultRunProperties38 = new A.DefaultRunProperties(){ FontSize = 1800, Kerning = 1200 };

            A.SolidFill solidFill26 = new A.SolidFill();
            A.SchemeColor schemeColor27 = new A.SchemeColor(){ Val = A.SchemeColorValues.Text1 };

            solidFill26.Append(schemeColor27);
            A.LatinFont latinFont23 = new A.LatinFont(){ Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont23 = new A.EastAsianFont(){ Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont23 = new A.ComplexScriptFont(){ Typeface = "+mn-cs" };

            defaultRunProperties38.Append(solidFill26);
            defaultRunProperties38.Append(latinFont23);
            defaultRunProperties38.Append(eastAsianFont23);
            defaultRunProperties38.Append(complexScriptFont23);

            level4ParagraphProperties4.Append(defaultRunProperties38);

            A.Level5ParagraphProperties level5ParagraphProperties4 = new A.Level5ParagraphProperties(){ LeftMargin = 1828800, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.DefaultRunProperties defaultRunProperties39 = new A.DefaultRunProperties(){ FontSize = 1800, Kerning = 1200 };

            A.SolidFill solidFill27 = new A.SolidFill();
            A.SchemeColor schemeColor28 = new A.SchemeColor(){ Val = A.SchemeColorValues.Text1 };

            solidFill27.Append(schemeColor28);
            A.LatinFont latinFont24 = new A.LatinFont(){ Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont24 = new A.EastAsianFont(){ Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont24 = new A.ComplexScriptFont(){ Typeface = "+mn-cs" };

            defaultRunProperties39.Append(solidFill27);
            defaultRunProperties39.Append(latinFont24);
            defaultRunProperties39.Append(eastAsianFont24);
            defaultRunProperties39.Append(complexScriptFont24);

            level5ParagraphProperties4.Append(defaultRunProperties39);

            A.Level6ParagraphProperties level6ParagraphProperties4 = new A.Level6ParagraphProperties(){ LeftMargin = 2286000, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.DefaultRunProperties defaultRunProperties40 = new A.DefaultRunProperties(){ FontSize = 1800, Kerning = 1200 };

            A.SolidFill solidFill28 = new A.SolidFill();
            A.SchemeColor schemeColor29 = new A.SchemeColor(){ Val = A.SchemeColorValues.Text1 };

            solidFill28.Append(schemeColor29);
            A.LatinFont latinFont25 = new A.LatinFont(){ Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont25 = new A.EastAsianFont(){ Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont25 = new A.ComplexScriptFont(){ Typeface = "+mn-cs" };

            defaultRunProperties40.Append(solidFill28);
            defaultRunProperties40.Append(latinFont25);
            defaultRunProperties40.Append(eastAsianFont25);
            defaultRunProperties40.Append(complexScriptFont25);

            level6ParagraphProperties4.Append(defaultRunProperties40);

            A.Level7ParagraphProperties level7ParagraphProperties4 = new A.Level7ParagraphProperties(){ LeftMargin = 2743200, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.DefaultRunProperties defaultRunProperties41 = new A.DefaultRunProperties(){ FontSize = 1800, Kerning = 1200 };

            A.SolidFill solidFill29 = new A.SolidFill();
            A.SchemeColor schemeColor30 = new A.SchemeColor(){ Val = A.SchemeColorValues.Text1 };

            solidFill29.Append(schemeColor30);
            A.LatinFont latinFont26 = new A.LatinFont(){ Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont26 = new A.EastAsianFont(){ Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont26 = new A.ComplexScriptFont(){ Typeface = "+mn-cs" };

            defaultRunProperties41.Append(solidFill29);
            defaultRunProperties41.Append(latinFont26);
            defaultRunProperties41.Append(eastAsianFont26);
            defaultRunProperties41.Append(complexScriptFont26);

            level7ParagraphProperties4.Append(defaultRunProperties41);

            A.Level8ParagraphProperties level8ParagraphProperties4 = new A.Level8ParagraphProperties(){ LeftMargin = 3200400, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.DefaultRunProperties defaultRunProperties42 = new A.DefaultRunProperties(){ FontSize = 1800, Kerning = 1200 };

            A.SolidFill solidFill30 = new A.SolidFill();
            A.SchemeColor schemeColor31 = new A.SchemeColor(){ Val = A.SchemeColorValues.Text1 };

            solidFill30.Append(schemeColor31);
            A.LatinFont latinFont27 = new A.LatinFont(){ Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont27 = new A.EastAsianFont(){ Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont27 = new A.ComplexScriptFont(){ Typeface = "+mn-cs" };

            defaultRunProperties42.Append(solidFill30);
            defaultRunProperties42.Append(latinFont27);
            defaultRunProperties42.Append(eastAsianFont27);
            defaultRunProperties42.Append(complexScriptFont27);

            level8ParagraphProperties4.Append(defaultRunProperties42);

            A.Level9ParagraphProperties level9ParagraphProperties4 = new A.Level9ParagraphProperties(){ LeftMargin = 3657600, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.DefaultRunProperties defaultRunProperties43 = new A.DefaultRunProperties(){ FontSize = 1800, Kerning = 1200 };

            A.SolidFill solidFill31 = new A.SolidFill();
            A.SchemeColor schemeColor32 = new A.SchemeColor(){ Val = A.SchemeColorValues.Text1 };

            solidFill31.Append(schemeColor32);
            A.LatinFont latinFont28 = new A.LatinFont(){ Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont28 = new A.EastAsianFont(){ Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont28 = new A.ComplexScriptFont(){ Typeface = "+mn-cs" };

            defaultRunProperties43.Append(solidFill31);
            defaultRunProperties43.Append(latinFont28);
            defaultRunProperties43.Append(eastAsianFont28);
            defaultRunProperties43.Append(complexScriptFont28);

            level9ParagraphProperties4.Append(defaultRunProperties43);

            otherStyle1.Append(defaultParagraphProperties2);
            otherStyle1.Append(level1ParagraphProperties9);
            otherStyle1.Append(level2ParagraphProperties4);
            otherStyle1.Append(level3ParagraphProperties4);
            otherStyle1.Append(level4ParagraphProperties4);
            otherStyle1.Append(level5ParagraphProperties4);
            otherStyle1.Append(level6ParagraphProperties4);
            otherStyle1.Append(level7ParagraphProperties4);
            otherStyle1.Append(level8ParagraphProperties4);
            otherStyle1.Append(level9ParagraphProperties4);

            textStyles1.Append(titleStyle1);
            textStyles1.Append(bodyStyle1);
            textStyles1.Append(otherStyle1);

            slideMaster1.Append(commonSlideData3);
            slideMaster1.Append(colorMap1);
            slideMaster1.Append(slideLayoutIdList1);
            slideMaster1.Append(textStyles1);

            slideMasterPart1.SlideMaster = slideMaster1;
        }

        // Generates content of slideLayoutPart2.
        private void GenerateSlideLayoutPart2Content(SlideLayoutPart slideLayoutPart2)
        {
            SlideLayout slideLayout2 = new SlideLayout(){ Type = SlideLayoutValues.ObjectText, Preserve = true };
            slideLayout2.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            slideLayout2.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            slideLayout2.AddNamespaceDeclaration("p", "http://schemas.openxmlformats.org/presentationml/2006/main");

            CommonSlideData commonSlideData4 = new CommonSlideData(){ Name = "Obsah s titulkem" };

            ShapeTree shapeTree4 = new ShapeTree();

            NonVisualGroupShapeProperties nonVisualGroupShapeProperties4 = new NonVisualGroupShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties16 = new NonVisualDrawingProperties(){ Id = (UInt32Value)1U, Name = "" };
            NonVisualGroupShapeDrawingProperties nonVisualGroupShapeDrawingProperties4 = new NonVisualGroupShapeDrawingProperties();
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties16 = new ApplicationNonVisualDrawingProperties();

            nonVisualGroupShapeProperties4.Append(nonVisualDrawingProperties16);
            nonVisualGroupShapeProperties4.Append(nonVisualGroupShapeDrawingProperties4);
            nonVisualGroupShapeProperties4.Append(applicationNonVisualDrawingProperties16);

            GroupShapeProperties groupShapeProperties4 = new GroupShapeProperties();

            A.TransformGroup transformGroup4 = new A.TransformGroup();
            A.Offset offset13 = new A.Offset(){ X = 0L, Y = 0L };
            A.Extents extents13 = new A.Extents(){ Cx = 0L, Cy = 0L };
            A.ChildOffset childOffset4 = new A.ChildOffset(){ X = 0L, Y = 0L };
            A.ChildExtents childExtents4 = new A.ChildExtents(){ Cx = 0L, Cy = 0L };

            transformGroup4.Append(offset13);
            transformGroup4.Append(extents13);
            transformGroup4.Append(childOffset4);
            transformGroup4.Append(childExtents4);

            groupShapeProperties4.Append(transformGroup4);

            Shape shape11 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties11 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties17 = new NonVisualDrawingProperties(){ Id = (UInt32Value)2U, Name = "Title 1" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties11 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks11 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties11.Append(shapeLocks11);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties17 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape11 = new PlaceholderShape(){ Type = PlaceholderValues.Title };

            applicationNonVisualDrawingProperties17.Append(placeholderShape11);

            nonVisualShapeProperties11.Append(nonVisualDrawingProperties17);
            nonVisualShapeProperties11.Append(nonVisualShapeDrawingProperties11);
            nonVisualShapeProperties11.Append(applicationNonVisualDrawingProperties17);

            ShapeProperties shapeProperties12 = new ShapeProperties();

            A.Transform2D transform2D9 = new A.Transform2D();
            A.Offset offset14 = new A.Offset(){ X = 629841L, Y = 457200L };
            A.Extents extents14 = new A.Extents(){ Cx = 2949178L, Cy = 1600200L };

            transform2D9.Append(offset14);
            transform2D9.Append(extents14);

            shapeProperties12.Append(transform2D9);

            TextBody textBody11 = new TextBody();
            A.BodyProperties bodyProperties11 = new A.BodyProperties(){ Anchor = A.TextAnchoringTypeValues.Bottom };

            A.ListStyle listStyle11 = new A.ListStyle();

            A.Level1ParagraphProperties level1ParagraphProperties10 = new A.Level1ParagraphProperties();
            A.DefaultRunProperties defaultRunProperties44 = new A.DefaultRunProperties(){ FontSize = 3200 };

            level1ParagraphProperties10.Append(defaultRunProperties44);

            listStyle11.Append(level1ParagraphProperties10);

            A.Paragraph paragraph15 = new A.Paragraph();

            A.Run run9 = new A.Run();
            A.RunProperties runProperties13 = new A.RunProperties(){ Language = "cs-CZ" };
            A.Text text13 = new A.Text();
            text13.Text = "Kliknutím lze upravit styl.";

            run9.Append(runProperties13);
            run9.Append(text13);
            A.EndParagraphRunProperties endParagraphRunProperties11 = new A.EndParagraphRunProperties(){ Language = "en-US", Dirty = false };

            paragraph15.Append(run9);
            paragraph15.Append(endParagraphRunProperties11);

            textBody11.Append(bodyProperties11);
            textBody11.Append(listStyle11);
            textBody11.Append(paragraph15);

            shape11.Append(nonVisualShapeProperties11);
            shape11.Append(shapeProperties12);
            shape11.Append(textBody11);

            Shape shape12 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties12 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties18 = new NonVisualDrawingProperties(){ Id = (UInt32Value)3U, Name = "Content Placeholder 2" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties12 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks12 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties12.Append(shapeLocks12);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties18 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape12 = new PlaceholderShape(){ Index = (UInt32Value)1U };

            applicationNonVisualDrawingProperties18.Append(placeholderShape12);

            nonVisualShapeProperties12.Append(nonVisualDrawingProperties18);
            nonVisualShapeProperties12.Append(nonVisualShapeDrawingProperties12);
            nonVisualShapeProperties12.Append(applicationNonVisualDrawingProperties18);

            ShapeProperties shapeProperties13 = new ShapeProperties();

            A.Transform2D transform2D10 = new A.Transform2D();
            A.Offset offset15 = new A.Offset(){ X = 3887391L, Y = 987426L };
            A.Extents extents15 = new A.Extents(){ Cx = 4629150L, Cy = 4873625L };

            transform2D10.Append(offset15);
            transform2D10.Append(extents15);

            shapeProperties13.Append(transform2D10);

            TextBody textBody12 = new TextBody();
            A.BodyProperties bodyProperties12 = new A.BodyProperties();

            A.ListStyle listStyle12 = new A.ListStyle();

            A.Level1ParagraphProperties level1ParagraphProperties11 = new A.Level1ParagraphProperties();
            A.DefaultRunProperties defaultRunProperties45 = new A.DefaultRunProperties(){ FontSize = 3200 };

            level1ParagraphProperties11.Append(defaultRunProperties45);

            A.Level2ParagraphProperties level2ParagraphProperties5 = new A.Level2ParagraphProperties();
            A.DefaultRunProperties defaultRunProperties46 = new A.DefaultRunProperties(){ FontSize = 2800 };

            level2ParagraphProperties5.Append(defaultRunProperties46);

            A.Level3ParagraphProperties level3ParagraphProperties5 = new A.Level3ParagraphProperties();
            A.DefaultRunProperties defaultRunProperties47 = new A.DefaultRunProperties(){ FontSize = 2400 };

            level3ParagraphProperties5.Append(defaultRunProperties47);

            A.Level4ParagraphProperties level4ParagraphProperties5 = new A.Level4ParagraphProperties();
            A.DefaultRunProperties defaultRunProperties48 = new A.DefaultRunProperties(){ FontSize = 2000 };

            level4ParagraphProperties5.Append(defaultRunProperties48);

            A.Level5ParagraphProperties level5ParagraphProperties5 = new A.Level5ParagraphProperties();
            A.DefaultRunProperties defaultRunProperties49 = new A.DefaultRunProperties(){ FontSize = 2000 };

            level5ParagraphProperties5.Append(defaultRunProperties49);

            A.Level6ParagraphProperties level6ParagraphProperties5 = new A.Level6ParagraphProperties();
            A.DefaultRunProperties defaultRunProperties50 = new A.DefaultRunProperties(){ FontSize = 2000 };

            level6ParagraphProperties5.Append(defaultRunProperties50);

            A.Level7ParagraphProperties level7ParagraphProperties5 = new A.Level7ParagraphProperties();
            A.DefaultRunProperties defaultRunProperties51 = new A.DefaultRunProperties(){ FontSize = 2000 };

            level7ParagraphProperties5.Append(defaultRunProperties51);

            A.Level8ParagraphProperties level8ParagraphProperties5 = new A.Level8ParagraphProperties();
            A.DefaultRunProperties defaultRunProperties52 = new A.DefaultRunProperties(){ FontSize = 2000 };

            level8ParagraphProperties5.Append(defaultRunProperties52);

            A.Level9ParagraphProperties level9ParagraphProperties5 = new A.Level9ParagraphProperties();
            A.DefaultRunProperties defaultRunProperties53 = new A.DefaultRunProperties(){ FontSize = 2000 };

            level9ParagraphProperties5.Append(defaultRunProperties53);

            listStyle12.Append(level1ParagraphProperties11);
            listStyle12.Append(level2ParagraphProperties5);
            listStyle12.Append(level3ParagraphProperties5);
            listStyle12.Append(level4ParagraphProperties5);
            listStyle12.Append(level5ParagraphProperties5);
            listStyle12.Append(level6ParagraphProperties5);
            listStyle12.Append(level7ParagraphProperties5);
            listStyle12.Append(level8ParagraphProperties5);
            listStyle12.Append(level9ParagraphProperties5);

            A.Paragraph paragraph16 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties6 = new A.ParagraphProperties(){ Level = 0 };

            A.Run run10 = new A.Run();
            A.RunProperties runProperties14 = new A.RunProperties(){ Language = "cs-CZ" };
            A.Text text14 = new A.Text();
            text14.Text = "Po kliknutí můžete upravovat styly textu v předloze.";

            run10.Append(runProperties14);
            run10.Append(text14);

            paragraph16.Append(paragraphProperties6);
            paragraph16.Append(run10);

            A.Paragraph paragraph17 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties7 = new A.ParagraphProperties(){ Level = 1 };

            A.Run run11 = new A.Run();
            A.RunProperties runProperties15 = new A.RunProperties(){ Language = "cs-CZ" };
            A.Text text15 = new A.Text();
            text15.Text = "Druhá úroveň";

            run11.Append(runProperties15);
            run11.Append(text15);

            paragraph17.Append(paragraphProperties7);
            paragraph17.Append(run11);

            A.Paragraph paragraph18 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties8 = new A.ParagraphProperties(){ Level = 2 };

            A.Run run12 = new A.Run();
            A.RunProperties runProperties16 = new A.RunProperties(){ Language = "cs-CZ" };
            A.Text text16 = new A.Text();
            text16.Text = "Třetí úroveň";

            run12.Append(runProperties16);
            run12.Append(text16);

            paragraph18.Append(paragraphProperties8);
            paragraph18.Append(run12);

            A.Paragraph paragraph19 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties9 = new A.ParagraphProperties(){ Level = 3 };

            A.Run run13 = new A.Run();
            A.RunProperties runProperties17 = new A.RunProperties(){ Language = "cs-CZ" };
            A.Text text17 = new A.Text();
            text17.Text = "Čtvrtá úroveň";

            run13.Append(runProperties17);
            run13.Append(text17);

            paragraph19.Append(paragraphProperties9);
            paragraph19.Append(run13);

            A.Paragraph paragraph20 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties10 = new A.ParagraphProperties(){ Level = 4 };

            A.Run run14 = new A.Run();
            A.RunProperties runProperties18 = new A.RunProperties(){ Language = "cs-CZ" };
            A.Text text18 = new A.Text();
            text18.Text = "Pátá úroveň";

            run14.Append(runProperties18);
            run14.Append(text18);
            A.EndParagraphRunProperties endParagraphRunProperties12 = new A.EndParagraphRunProperties(){ Language = "en-US", Dirty = false };

            paragraph20.Append(paragraphProperties10);
            paragraph20.Append(run14);
            paragraph20.Append(endParagraphRunProperties12);

            textBody12.Append(bodyProperties12);
            textBody12.Append(listStyle12);
            textBody12.Append(paragraph16);
            textBody12.Append(paragraph17);
            textBody12.Append(paragraph18);
            textBody12.Append(paragraph19);
            textBody12.Append(paragraph20);

            shape12.Append(nonVisualShapeProperties12);
            shape12.Append(shapeProperties13);
            shape12.Append(textBody12);

            Shape shape13 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties13 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties19 = new NonVisualDrawingProperties(){ Id = (UInt32Value)4U, Name = "Text Placeholder 3" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties13 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks13 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties13.Append(shapeLocks13);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties19 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape13 = new PlaceholderShape(){ Type = PlaceholderValues.Body, Size = PlaceholderSizeValues.Half, Index = (UInt32Value)2U };

            applicationNonVisualDrawingProperties19.Append(placeholderShape13);

            nonVisualShapeProperties13.Append(nonVisualDrawingProperties19);
            nonVisualShapeProperties13.Append(nonVisualShapeDrawingProperties13);
            nonVisualShapeProperties13.Append(applicationNonVisualDrawingProperties19);

            ShapeProperties shapeProperties14 = new ShapeProperties();

            A.Transform2D transform2D11 = new A.Transform2D();
            A.Offset offset16 = new A.Offset(){ X = 629841L, Y = 2057400L };
            A.Extents extents16 = new A.Extents(){ Cx = 2949178L, Cy = 3811588L };

            transform2D11.Append(offset16);
            transform2D11.Append(extents16);

            shapeProperties14.Append(transform2D11);

            TextBody textBody13 = new TextBody();
            A.BodyProperties bodyProperties13 = new A.BodyProperties();

            A.ListStyle listStyle13 = new A.ListStyle();

            A.Level1ParagraphProperties level1ParagraphProperties12 = new A.Level1ParagraphProperties(){ LeftMargin = 0, Indent = 0 };
            A.NoBullet noBullet11 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties54 = new A.DefaultRunProperties(){ FontSize = 1600 };

            level1ParagraphProperties12.Append(noBullet11);
            level1ParagraphProperties12.Append(defaultRunProperties54);

            A.Level2ParagraphProperties level2ParagraphProperties6 = new A.Level2ParagraphProperties(){ LeftMargin = 457200, Indent = 0 };
            A.NoBullet noBullet12 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties55 = new A.DefaultRunProperties(){ FontSize = 1400 };

            level2ParagraphProperties6.Append(noBullet12);
            level2ParagraphProperties6.Append(defaultRunProperties55);

            A.Level3ParagraphProperties level3ParagraphProperties6 = new A.Level3ParagraphProperties(){ LeftMargin = 914400, Indent = 0 };
            A.NoBullet noBullet13 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties56 = new A.DefaultRunProperties(){ FontSize = 1200 };

            level3ParagraphProperties6.Append(noBullet13);
            level3ParagraphProperties6.Append(defaultRunProperties56);

            A.Level4ParagraphProperties level4ParagraphProperties6 = new A.Level4ParagraphProperties(){ LeftMargin = 1371600, Indent = 0 };
            A.NoBullet noBullet14 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties57 = new A.DefaultRunProperties(){ FontSize = 1000 };

            level4ParagraphProperties6.Append(noBullet14);
            level4ParagraphProperties6.Append(defaultRunProperties57);

            A.Level5ParagraphProperties level5ParagraphProperties6 = new A.Level5ParagraphProperties(){ LeftMargin = 1828800, Indent = 0 };
            A.NoBullet noBullet15 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties58 = new A.DefaultRunProperties(){ FontSize = 1000 };

            level5ParagraphProperties6.Append(noBullet15);
            level5ParagraphProperties6.Append(defaultRunProperties58);

            A.Level6ParagraphProperties level6ParagraphProperties6 = new A.Level6ParagraphProperties(){ LeftMargin = 2286000, Indent = 0 };
            A.NoBullet noBullet16 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties59 = new A.DefaultRunProperties(){ FontSize = 1000 };

            level6ParagraphProperties6.Append(noBullet16);
            level6ParagraphProperties6.Append(defaultRunProperties59);

            A.Level7ParagraphProperties level7ParagraphProperties6 = new A.Level7ParagraphProperties(){ LeftMargin = 2743200, Indent = 0 };
            A.NoBullet noBullet17 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties60 = new A.DefaultRunProperties(){ FontSize = 1000 };

            level7ParagraphProperties6.Append(noBullet17);
            level7ParagraphProperties6.Append(defaultRunProperties60);

            A.Level8ParagraphProperties level8ParagraphProperties6 = new A.Level8ParagraphProperties(){ LeftMargin = 3200400, Indent = 0 };
            A.NoBullet noBullet18 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties61 = new A.DefaultRunProperties(){ FontSize = 1000 };

            level8ParagraphProperties6.Append(noBullet18);
            level8ParagraphProperties6.Append(defaultRunProperties61);

            A.Level9ParagraphProperties level9ParagraphProperties6 = new A.Level9ParagraphProperties(){ LeftMargin = 3657600, Indent = 0 };
            A.NoBullet noBullet19 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties62 = new A.DefaultRunProperties(){ FontSize = 1000 };

            level9ParagraphProperties6.Append(noBullet19);
            level9ParagraphProperties6.Append(defaultRunProperties62);

            listStyle13.Append(level1ParagraphProperties12);
            listStyle13.Append(level2ParagraphProperties6);
            listStyle13.Append(level3ParagraphProperties6);
            listStyle13.Append(level4ParagraphProperties6);
            listStyle13.Append(level5ParagraphProperties6);
            listStyle13.Append(level6ParagraphProperties6);
            listStyle13.Append(level7ParagraphProperties6);
            listStyle13.Append(level8ParagraphProperties6);
            listStyle13.Append(level9ParagraphProperties6);

            A.Paragraph paragraph21 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties11 = new A.ParagraphProperties(){ Level = 0 };

            A.Run run15 = new A.Run();
            A.RunProperties runProperties19 = new A.RunProperties(){ Language = "cs-CZ" };
            A.Text text19 = new A.Text();
            text19.Text = "Po kliknutí můžete upravovat styly textu v předloze.";

            run15.Append(runProperties19);
            run15.Append(text19);

            paragraph21.Append(paragraphProperties11);
            paragraph21.Append(run15);

            textBody13.Append(bodyProperties13);
            textBody13.Append(listStyle13);
            textBody13.Append(paragraph21);

            shape13.Append(nonVisualShapeProperties13);
            shape13.Append(shapeProperties14);
            shape13.Append(textBody13);

            Shape shape14 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties14 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties20 = new NonVisualDrawingProperties(){ Id = (UInt32Value)5U, Name = "Date Placeholder 4" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties14 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks14 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties14.Append(shapeLocks14);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties20 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape14 = new PlaceholderShape(){ Type = PlaceholderValues.DateAndTime, Size = PlaceholderSizeValues.Half, Index = (UInt32Value)10U };

            applicationNonVisualDrawingProperties20.Append(placeholderShape14);

            nonVisualShapeProperties14.Append(nonVisualDrawingProperties20);
            nonVisualShapeProperties14.Append(nonVisualShapeDrawingProperties14);
            nonVisualShapeProperties14.Append(applicationNonVisualDrawingProperties20);
            ShapeProperties shapeProperties15 = new ShapeProperties();

            TextBody textBody14 = new TextBody();
            A.BodyProperties bodyProperties14 = new A.BodyProperties();
            A.ListStyle listStyle14 = new A.ListStyle();

            A.Paragraph paragraph22 = new A.Paragraph();

            A.Field field5 = new A.Field(){ Id = "{1D065981-0097-4BA9-A692-9FF652AFAF35}", Type = "datetimeFigureOut" };

            A.RunProperties runProperties20 = new A.RunProperties(){ Language = "cs-CZ" };
            runProperties20.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            A.Text text20 = new A.Text();
            text20.Text = "14.03.2023";

            field5.Append(runProperties20);
            field5.Append(text20);
            A.EndParagraphRunProperties endParagraphRunProperties13 = new A.EndParagraphRunProperties(){ Language = "cs-CZ" };

            paragraph22.Append(field5);
            paragraph22.Append(endParagraphRunProperties13);

            textBody14.Append(bodyProperties14);
            textBody14.Append(listStyle14);
            textBody14.Append(paragraph22);

            shape14.Append(nonVisualShapeProperties14);
            shape14.Append(shapeProperties15);
            shape14.Append(textBody14);

            Shape shape15 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties15 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties21 = new NonVisualDrawingProperties(){ Id = (UInt32Value)6U, Name = "Footer Placeholder 5" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties15 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks15 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties15.Append(shapeLocks15);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties21 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape15 = new PlaceholderShape(){ Type = PlaceholderValues.Footer, Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)11U };

            applicationNonVisualDrawingProperties21.Append(placeholderShape15);

            nonVisualShapeProperties15.Append(nonVisualDrawingProperties21);
            nonVisualShapeProperties15.Append(nonVisualShapeDrawingProperties15);
            nonVisualShapeProperties15.Append(applicationNonVisualDrawingProperties21);
            ShapeProperties shapeProperties16 = new ShapeProperties();

            TextBody textBody15 = new TextBody();
            A.BodyProperties bodyProperties15 = new A.BodyProperties();
            A.ListStyle listStyle15 = new A.ListStyle();

            A.Paragraph paragraph23 = new A.Paragraph();
            A.EndParagraphRunProperties endParagraphRunProperties14 = new A.EndParagraphRunProperties(){ Language = "cs-CZ" };

            paragraph23.Append(endParagraphRunProperties14);

            textBody15.Append(bodyProperties15);
            textBody15.Append(listStyle15);
            textBody15.Append(paragraph23);

            shape15.Append(nonVisualShapeProperties15);
            shape15.Append(shapeProperties16);
            shape15.Append(textBody15);

            Shape shape16 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties16 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties22 = new NonVisualDrawingProperties(){ Id = (UInt32Value)7U, Name = "Slide Number Placeholder 6" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties16 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks16 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties16.Append(shapeLocks16);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties22 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape16 = new PlaceholderShape(){ Type = PlaceholderValues.SlideNumber, Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)12U };

            applicationNonVisualDrawingProperties22.Append(placeholderShape16);

            nonVisualShapeProperties16.Append(nonVisualDrawingProperties22);
            nonVisualShapeProperties16.Append(nonVisualShapeDrawingProperties16);
            nonVisualShapeProperties16.Append(applicationNonVisualDrawingProperties22);
            ShapeProperties shapeProperties17 = new ShapeProperties();

            TextBody textBody16 = new TextBody();
            A.BodyProperties bodyProperties16 = new A.BodyProperties();
            A.ListStyle listStyle16 = new A.ListStyle();

            A.Paragraph paragraph24 = new A.Paragraph();

            A.Field field6 = new A.Field(){ Id = "{5D7F2B33-3E35-4F97-A448-9FEB151DF743}", Type = "slidenum" };

            A.RunProperties runProperties21 = new A.RunProperties(){ Language = "cs-CZ" };
            runProperties21.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            A.Text text21 = new A.Text();
            text21.Text = "‹#›";

            field6.Append(runProperties21);
            field6.Append(text21);
            A.EndParagraphRunProperties endParagraphRunProperties15 = new A.EndParagraphRunProperties(){ Language = "cs-CZ" };

            paragraph24.Append(field6);
            paragraph24.Append(endParagraphRunProperties15);

            textBody16.Append(bodyProperties16);
            textBody16.Append(listStyle16);
            textBody16.Append(paragraph24);

            shape16.Append(nonVisualShapeProperties16);
            shape16.Append(shapeProperties17);
            shape16.Append(textBody16);

            shapeTree4.Append(nonVisualGroupShapeProperties4);
            shapeTree4.Append(groupShapeProperties4);
            shapeTree4.Append(shape11);
            shapeTree4.Append(shape12);
            shapeTree4.Append(shape13);
            shapeTree4.Append(shape14);
            shapeTree4.Append(shape15);
            shapeTree4.Append(shape16);

            CommonSlideDataExtensionList commonSlideDataExtensionList4 = new CommonSlideDataExtensionList();

            CommonSlideDataExtension commonSlideDataExtension4 = new CommonSlideDataExtension(){ Uri = "{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}" };

            P14.CreationId creationId4 = new P14.CreationId(){ Val = (UInt32Value)2967877822U };
            creationId4.AddNamespaceDeclaration("p14", "http://schemas.microsoft.com/office/powerpoint/2010/main");

            commonSlideDataExtension4.Append(creationId4);

            commonSlideDataExtensionList4.Append(commonSlideDataExtension4);

            commonSlideData4.Append(shapeTree4);
            commonSlideData4.Append(commonSlideDataExtensionList4);

            ColorMapOverride colorMapOverride3 = new ColorMapOverride();
            A.MasterColorMapping masterColorMapping3 = new A.MasterColorMapping();

            colorMapOverride3.Append(masterColorMapping3);

            slideLayout2.Append(commonSlideData4);
            slideLayout2.Append(colorMapOverride3);

            slideLayoutPart2.SlideLayout = slideLayout2;
        }

        // Generates content of slideLayoutPart3.
        private void GenerateSlideLayoutPart3Content(SlideLayoutPart slideLayoutPart3)
        {
            SlideLayout slideLayout3 = new SlideLayout(){ Type = SlideLayoutValues.SectionHeader, Preserve = true };
            slideLayout3.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            slideLayout3.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            slideLayout3.AddNamespaceDeclaration("p", "http://schemas.openxmlformats.org/presentationml/2006/main");

            CommonSlideData commonSlideData5 = new CommonSlideData(){ Name = "Záhlaví oddílu" };

            ShapeTree shapeTree5 = new ShapeTree();

            NonVisualGroupShapeProperties nonVisualGroupShapeProperties5 = new NonVisualGroupShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties23 = new NonVisualDrawingProperties(){ Id = (UInt32Value)1U, Name = "" };
            NonVisualGroupShapeDrawingProperties nonVisualGroupShapeDrawingProperties5 = new NonVisualGroupShapeDrawingProperties();
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties23 = new ApplicationNonVisualDrawingProperties();

            nonVisualGroupShapeProperties5.Append(nonVisualDrawingProperties23);
            nonVisualGroupShapeProperties5.Append(nonVisualGroupShapeDrawingProperties5);
            nonVisualGroupShapeProperties5.Append(applicationNonVisualDrawingProperties23);

            GroupShapeProperties groupShapeProperties5 = new GroupShapeProperties();

            A.TransformGroup transformGroup5 = new A.TransformGroup();
            A.Offset offset17 = new A.Offset(){ X = 0L, Y = 0L };
            A.Extents extents17 = new A.Extents(){ Cx = 0L, Cy = 0L };
            A.ChildOffset childOffset5 = new A.ChildOffset(){ X = 0L, Y = 0L };
            A.ChildExtents childExtents5 = new A.ChildExtents(){ Cx = 0L, Cy = 0L };

            transformGroup5.Append(offset17);
            transformGroup5.Append(extents17);
            transformGroup5.Append(childOffset5);
            transformGroup5.Append(childExtents5);

            groupShapeProperties5.Append(transformGroup5);

            Shape shape17 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties17 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties24 = new NonVisualDrawingProperties(){ Id = (UInt32Value)2U, Name = "Title 1" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties17 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks17 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties17.Append(shapeLocks17);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties24 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape17 = new PlaceholderShape(){ Type = PlaceholderValues.Title };

            applicationNonVisualDrawingProperties24.Append(placeholderShape17);

            nonVisualShapeProperties17.Append(nonVisualDrawingProperties24);
            nonVisualShapeProperties17.Append(nonVisualShapeDrawingProperties17);
            nonVisualShapeProperties17.Append(applicationNonVisualDrawingProperties24);

            ShapeProperties shapeProperties18 = new ShapeProperties();

            A.Transform2D transform2D12 = new A.Transform2D();
            A.Offset offset18 = new A.Offset(){ X = 623888L, Y = 1709739L };
            A.Extents extents18 = new A.Extents(){ Cx = 7886700L, Cy = 2852737L };

            transform2D12.Append(offset18);
            transform2D12.Append(extents18);

            shapeProperties18.Append(transform2D12);

            TextBody textBody17 = new TextBody();
            A.BodyProperties bodyProperties17 = new A.BodyProperties(){ Anchor = A.TextAnchoringTypeValues.Bottom };

            A.ListStyle listStyle17 = new A.ListStyle();

            A.Level1ParagraphProperties level1ParagraphProperties13 = new A.Level1ParagraphProperties();
            A.DefaultRunProperties defaultRunProperties63 = new A.DefaultRunProperties(){ FontSize = 6000 };

            level1ParagraphProperties13.Append(defaultRunProperties63);

            listStyle17.Append(level1ParagraphProperties13);

            A.Paragraph paragraph25 = new A.Paragraph();

            A.Run run16 = new A.Run();
            A.RunProperties runProperties22 = new A.RunProperties(){ Language = "cs-CZ" };
            A.Text text22 = new A.Text();
            text22.Text = "Kliknutím lze upravit styl.";

            run16.Append(runProperties22);
            run16.Append(text22);
            A.EndParagraphRunProperties endParagraphRunProperties16 = new A.EndParagraphRunProperties(){ Language = "en-US", Dirty = false };

            paragraph25.Append(run16);
            paragraph25.Append(endParagraphRunProperties16);

            textBody17.Append(bodyProperties17);
            textBody17.Append(listStyle17);
            textBody17.Append(paragraph25);

            shape17.Append(nonVisualShapeProperties17);
            shape17.Append(shapeProperties18);
            shape17.Append(textBody17);

            Shape shape18 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties18 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties25 = new NonVisualDrawingProperties(){ Id = (UInt32Value)3U, Name = "Text Placeholder 2" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties18 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks18 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties18.Append(shapeLocks18);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties25 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape18 = new PlaceholderShape(){ Type = PlaceholderValues.Body, Index = (UInt32Value)1U };

            applicationNonVisualDrawingProperties25.Append(placeholderShape18);

            nonVisualShapeProperties18.Append(nonVisualDrawingProperties25);
            nonVisualShapeProperties18.Append(nonVisualShapeDrawingProperties18);
            nonVisualShapeProperties18.Append(applicationNonVisualDrawingProperties25);

            ShapeProperties shapeProperties19 = new ShapeProperties();

            A.Transform2D transform2D13 = new A.Transform2D();
            A.Offset offset19 = new A.Offset(){ X = 623888L, Y = 4589464L };
            A.Extents extents19 = new A.Extents(){ Cx = 7886700L, Cy = 1500187L };

            transform2D13.Append(offset19);
            transform2D13.Append(extents19);

            shapeProperties19.Append(transform2D13);

            TextBody textBody18 = new TextBody();
            A.BodyProperties bodyProperties18 = new A.BodyProperties();

            A.ListStyle listStyle18 = new A.ListStyle();

            A.Level1ParagraphProperties level1ParagraphProperties14 = new A.Level1ParagraphProperties(){ LeftMargin = 0, Indent = 0 };
            A.NoBullet noBullet20 = new A.NoBullet();

            A.DefaultRunProperties defaultRunProperties64 = new A.DefaultRunProperties(){ FontSize = 2400 };

            A.SolidFill solidFill32 = new A.SolidFill();
            A.SchemeColor schemeColor33 = new A.SchemeColor(){ Val = A.SchemeColorValues.Text1 };

            solidFill32.Append(schemeColor33);

            defaultRunProperties64.Append(solidFill32);

            level1ParagraphProperties14.Append(noBullet20);
            level1ParagraphProperties14.Append(defaultRunProperties64);

            A.Level2ParagraphProperties level2ParagraphProperties7 = new A.Level2ParagraphProperties(){ LeftMargin = 457200, Indent = 0 };
            A.NoBullet noBullet21 = new A.NoBullet();

            A.DefaultRunProperties defaultRunProperties65 = new A.DefaultRunProperties(){ FontSize = 2000 };

            A.SolidFill solidFill33 = new A.SolidFill();

            A.SchemeColor schemeColor34 = new A.SchemeColor(){ Val = A.SchemeColorValues.Text1 };
            A.Tint tint4 = new A.Tint(){ Val = 75000 };

            schemeColor34.Append(tint4);

            solidFill33.Append(schemeColor34);

            defaultRunProperties65.Append(solidFill33);

            level2ParagraphProperties7.Append(noBullet21);
            level2ParagraphProperties7.Append(defaultRunProperties65);

            A.Level3ParagraphProperties level3ParagraphProperties7 = new A.Level3ParagraphProperties(){ LeftMargin = 914400, Indent = 0 };
            A.NoBullet noBullet22 = new A.NoBullet();

            A.DefaultRunProperties defaultRunProperties66 = new A.DefaultRunProperties(){ FontSize = 1800 };

            A.SolidFill solidFill34 = new A.SolidFill();

            A.SchemeColor schemeColor35 = new A.SchemeColor(){ Val = A.SchemeColorValues.Text1 };
            A.Tint tint5 = new A.Tint(){ Val = 75000 };

            schemeColor35.Append(tint5);

            solidFill34.Append(schemeColor35);

            defaultRunProperties66.Append(solidFill34);

            level3ParagraphProperties7.Append(noBullet22);
            level3ParagraphProperties7.Append(defaultRunProperties66);

            A.Level4ParagraphProperties level4ParagraphProperties7 = new A.Level4ParagraphProperties(){ LeftMargin = 1371600, Indent = 0 };
            A.NoBullet noBullet23 = new A.NoBullet();

            A.DefaultRunProperties defaultRunProperties67 = new A.DefaultRunProperties(){ FontSize = 1600 };

            A.SolidFill solidFill35 = new A.SolidFill();

            A.SchemeColor schemeColor36 = new A.SchemeColor(){ Val = A.SchemeColorValues.Text1 };
            A.Tint tint6 = new A.Tint(){ Val = 75000 };

            schemeColor36.Append(tint6);

            solidFill35.Append(schemeColor36);

            defaultRunProperties67.Append(solidFill35);

            level4ParagraphProperties7.Append(noBullet23);
            level4ParagraphProperties7.Append(defaultRunProperties67);

            A.Level5ParagraphProperties level5ParagraphProperties7 = new A.Level5ParagraphProperties(){ LeftMargin = 1828800, Indent = 0 };
            A.NoBullet noBullet24 = new A.NoBullet();

            A.DefaultRunProperties defaultRunProperties68 = new A.DefaultRunProperties(){ FontSize = 1600 };

            A.SolidFill solidFill36 = new A.SolidFill();

            A.SchemeColor schemeColor37 = new A.SchemeColor(){ Val = A.SchemeColorValues.Text1 };
            A.Tint tint7 = new A.Tint(){ Val = 75000 };

            schemeColor37.Append(tint7);

            solidFill36.Append(schemeColor37);

            defaultRunProperties68.Append(solidFill36);

            level5ParagraphProperties7.Append(noBullet24);
            level5ParagraphProperties7.Append(defaultRunProperties68);

            A.Level6ParagraphProperties level6ParagraphProperties7 = new A.Level6ParagraphProperties(){ LeftMargin = 2286000, Indent = 0 };
            A.NoBullet noBullet25 = new A.NoBullet();

            A.DefaultRunProperties defaultRunProperties69 = new A.DefaultRunProperties(){ FontSize = 1600 };

            A.SolidFill solidFill37 = new A.SolidFill();

            A.SchemeColor schemeColor38 = new A.SchemeColor(){ Val = A.SchemeColorValues.Text1 };
            A.Tint tint8 = new A.Tint(){ Val = 75000 };

            schemeColor38.Append(tint8);

            solidFill37.Append(schemeColor38);

            defaultRunProperties69.Append(solidFill37);

            level6ParagraphProperties7.Append(noBullet25);
            level6ParagraphProperties7.Append(defaultRunProperties69);

            A.Level7ParagraphProperties level7ParagraphProperties7 = new A.Level7ParagraphProperties(){ LeftMargin = 2743200, Indent = 0 };
            A.NoBullet noBullet26 = new A.NoBullet();

            A.DefaultRunProperties defaultRunProperties70 = new A.DefaultRunProperties(){ FontSize = 1600 };

            A.SolidFill solidFill38 = new A.SolidFill();

            A.SchemeColor schemeColor39 = new A.SchemeColor(){ Val = A.SchemeColorValues.Text1 };
            A.Tint tint9 = new A.Tint(){ Val = 75000 };

            schemeColor39.Append(tint9);

            solidFill38.Append(schemeColor39);

            defaultRunProperties70.Append(solidFill38);

            level7ParagraphProperties7.Append(noBullet26);
            level7ParagraphProperties7.Append(defaultRunProperties70);

            A.Level8ParagraphProperties level8ParagraphProperties7 = new A.Level8ParagraphProperties(){ LeftMargin = 3200400, Indent = 0 };
            A.NoBullet noBullet27 = new A.NoBullet();

            A.DefaultRunProperties defaultRunProperties71 = new A.DefaultRunProperties(){ FontSize = 1600 };

            A.SolidFill solidFill39 = new A.SolidFill();

            A.SchemeColor schemeColor40 = new A.SchemeColor(){ Val = A.SchemeColorValues.Text1 };
            A.Tint tint10 = new A.Tint(){ Val = 75000 };

            schemeColor40.Append(tint10);

            solidFill39.Append(schemeColor40);

            defaultRunProperties71.Append(solidFill39);

            level8ParagraphProperties7.Append(noBullet27);
            level8ParagraphProperties7.Append(defaultRunProperties71);

            A.Level9ParagraphProperties level9ParagraphProperties7 = new A.Level9ParagraphProperties(){ LeftMargin = 3657600, Indent = 0 };
            A.NoBullet noBullet28 = new A.NoBullet();

            A.DefaultRunProperties defaultRunProperties72 = new A.DefaultRunProperties(){ FontSize = 1600 };

            A.SolidFill solidFill40 = new A.SolidFill();

            A.SchemeColor schemeColor41 = new A.SchemeColor(){ Val = A.SchemeColorValues.Text1 };
            A.Tint tint11 = new A.Tint(){ Val = 75000 };

            schemeColor41.Append(tint11);

            solidFill40.Append(schemeColor41);

            defaultRunProperties72.Append(solidFill40);

            level9ParagraphProperties7.Append(noBullet28);
            level9ParagraphProperties7.Append(defaultRunProperties72);

            listStyle18.Append(level1ParagraphProperties14);
            listStyle18.Append(level2ParagraphProperties7);
            listStyle18.Append(level3ParagraphProperties7);
            listStyle18.Append(level4ParagraphProperties7);
            listStyle18.Append(level5ParagraphProperties7);
            listStyle18.Append(level6ParagraphProperties7);
            listStyle18.Append(level7ParagraphProperties7);
            listStyle18.Append(level8ParagraphProperties7);
            listStyle18.Append(level9ParagraphProperties7);

            A.Paragraph paragraph26 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties12 = new A.ParagraphProperties(){ Level = 0 };

            A.Run run17 = new A.Run();
            A.RunProperties runProperties23 = new A.RunProperties(){ Language = "cs-CZ" };
            A.Text text23 = new A.Text();
            text23.Text = "Po kliknutí můžete upravovat styly textu v předloze.";

            run17.Append(runProperties23);
            run17.Append(text23);

            paragraph26.Append(paragraphProperties12);
            paragraph26.Append(run17);

            textBody18.Append(bodyProperties18);
            textBody18.Append(listStyle18);
            textBody18.Append(paragraph26);

            shape18.Append(nonVisualShapeProperties18);
            shape18.Append(shapeProperties19);
            shape18.Append(textBody18);

            Shape shape19 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties19 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties26 = new NonVisualDrawingProperties(){ Id = (UInt32Value)4U, Name = "Date Placeholder 3" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties19 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks19 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties19.Append(shapeLocks19);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties26 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape19 = new PlaceholderShape(){ Type = PlaceholderValues.DateAndTime, Size = PlaceholderSizeValues.Half, Index = (UInt32Value)10U };

            applicationNonVisualDrawingProperties26.Append(placeholderShape19);

            nonVisualShapeProperties19.Append(nonVisualDrawingProperties26);
            nonVisualShapeProperties19.Append(nonVisualShapeDrawingProperties19);
            nonVisualShapeProperties19.Append(applicationNonVisualDrawingProperties26);
            ShapeProperties shapeProperties20 = new ShapeProperties();

            TextBody textBody19 = new TextBody();
            A.BodyProperties bodyProperties19 = new A.BodyProperties();
            A.ListStyle listStyle19 = new A.ListStyle();

            A.Paragraph paragraph27 = new A.Paragraph();

            A.Field field7 = new A.Field(){ Id = "{1D065981-0097-4BA9-A692-9FF652AFAF35}", Type = "datetimeFigureOut" };

            A.RunProperties runProperties24 = new A.RunProperties(){ Language = "cs-CZ" };
            runProperties24.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            A.Text text24 = new A.Text();
            text24.Text = "14.03.2023";

            field7.Append(runProperties24);
            field7.Append(text24);
            A.EndParagraphRunProperties endParagraphRunProperties17 = new A.EndParagraphRunProperties(){ Language = "cs-CZ" };

            paragraph27.Append(field7);
            paragraph27.Append(endParagraphRunProperties17);

            textBody19.Append(bodyProperties19);
            textBody19.Append(listStyle19);
            textBody19.Append(paragraph27);

            shape19.Append(nonVisualShapeProperties19);
            shape19.Append(shapeProperties20);
            shape19.Append(textBody19);

            Shape shape20 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties20 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties27 = new NonVisualDrawingProperties(){ Id = (UInt32Value)5U, Name = "Footer Placeholder 4" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties20 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks20 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties20.Append(shapeLocks20);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties27 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape20 = new PlaceholderShape(){ Type = PlaceholderValues.Footer, Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)11U };

            applicationNonVisualDrawingProperties27.Append(placeholderShape20);

            nonVisualShapeProperties20.Append(nonVisualDrawingProperties27);
            nonVisualShapeProperties20.Append(nonVisualShapeDrawingProperties20);
            nonVisualShapeProperties20.Append(applicationNonVisualDrawingProperties27);
            ShapeProperties shapeProperties21 = new ShapeProperties();

            TextBody textBody20 = new TextBody();
            A.BodyProperties bodyProperties20 = new A.BodyProperties();
            A.ListStyle listStyle20 = new A.ListStyle();

            A.Paragraph paragraph28 = new A.Paragraph();
            A.EndParagraphRunProperties endParagraphRunProperties18 = new A.EndParagraphRunProperties(){ Language = "cs-CZ" };

            paragraph28.Append(endParagraphRunProperties18);

            textBody20.Append(bodyProperties20);
            textBody20.Append(listStyle20);
            textBody20.Append(paragraph28);

            shape20.Append(nonVisualShapeProperties20);
            shape20.Append(shapeProperties21);
            shape20.Append(textBody20);

            Shape shape21 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties21 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties28 = new NonVisualDrawingProperties(){ Id = (UInt32Value)6U, Name = "Slide Number Placeholder 5" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties21 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks21 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties21.Append(shapeLocks21);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties28 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape21 = new PlaceholderShape(){ Type = PlaceholderValues.SlideNumber, Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)12U };

            applicationNonVisualDrawingProperties28.Append(placeholderShape21);

            nonVisualShapeProperties21.Append(nonVisualDrawingProperties28);
            nonVisualShapeProperties21.Append(nonVisualShapeDrawingProperties21);
            nonVisualShapeProperties21.Append(applicationNonVisualDrawingProperties28);
            ShapeProperties shapeProperties22 = new ShapeProperties();

            TextBody textBody21 = new TextBody();
            A.BodyProperties bodyProperties21 = new A.BodyProperties();
            A.ListStyle listStyle21 = new A.ListStyle();

            A.Paragraph paragraph29 = new A.Paragraph();

            A.Field field8 = new A.Field(){ Id = "{5D7F2B33-3E35-4F97-A448-9FEB151DF743}", Type = "slidenum" };

            A.RunProperties runProperties25 = new A.RunProperties(){ Language = "cs-CZ" };
            runProperties25.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            A.Text text25 = new A.Text();
            text25.Text = "‹#›";

            field8.Append(runProperties25);
            field8.Append(text25);
            A.EndParagraphRunProperties endParagraphRunProperties19 = new A.EndParagraphRunProperties(){ Language = "cs-CZ" };

            paragraph29.Append(field8);
            paragraph29.Append(endParagraphRunProperties19);

            textBody21.Append(bodyProperties21);
            textBody21.Append(listStyle21);
            textBody21.Append(paragraph29);

            shape21.Append(nonVisualShapeProperties21);
            shape21.Append(shapeProperties22);
            shape21.Append(textBody21);

            shapeTree5.Append(nonVisualGroupShapeProperties5);
            shapeTree5.Append(groupShapeProperties5);
            shapeTree5.Append(shape17);
            shapeTree5.Append(shape18);
            shapeTree5.Append(shape19);
            shapeTree5.Append(shape20);
            shapeTree5.Append(shape21);

            CommonSlideDataExtensionList commonSlideDataExtensionList5 = new CommonSlideDataExtensionList();

            CommonSlideDataExtension commonSlideDataExtension5 = new CommonSlideDataExtension(){ Uri = "{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}" };

            P14.CreationId creationId5 = new P14.CreationId(){ Val = (UInt32Value)569065252U };
            creationId5.AddNamespaceDeclaration("p14", "http://schemas.microsoft.com/office/powerpoint/2010/main");

            commonSlideDataExtension5.Append(creationId5);

            commonSlideDataExtensionList5.Append(commonSlideDataExtension5);

            commonSlideData5.Append(shapeTree5);
            commonSlideData5.Append(commonSlideDataExtensionList5);

            ColorMapOverride colorMapOverride4 = new ColorMapOverride();
            A.MasterColorMapping masterColorMapping4 = new A.MasterColorMapping();

            colorMapOverride4.Append(masterColorMapping4);

            slideLayout3.Append(commonSlideData5);
            slideLayout3.Append(colorMapOverride4);

            slideLayoutPart3.SlideLayout = slideLayout3;
        }

        // Generates content of slideLayoutPart4.
        private void GenerateSlideLayoutPart4Content(SlideLayoutPart slideLayoutPart4)
        {
            SlideLayout slideLayout4 = new SlideLayout(){ Type = SlideLayoutValues.Blank, Preserve = true };
            slideLayout4.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            slideLayout4.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            slideLayout4.AddNamespaceDeclaration("p", "http://schemas.openxmlformats.org/presentationml/2006/main");

            CommonSlideData commonSlideData6 = new CommonSlideData(){ Name = "Prázdný" };

            ShapeTree shapeTree6 = new ShapeTree();

            NonVisualGroupShapeProperties nonVisualGroupShapeProperties6 = new NonVisualGroupShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties29 = new NonVisualDrawingProperties(){ Id = (UInt32Value)1U, Name = "" };
            NonVisualGroupShapeDrawingProperties nonVisualGroupShapeDrawingProperties6 = new NonVisualGroupShapeDrawingProperties();
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties29 = new ApplicationNonVisualDrawingProperties();

            nonVisualGroupShapeProperties6.Append(nonVisualDrawingProperties29);
            nonVisualGroupShapeProperties6.Append(nonVisualGroupShapeDrawingProperties6);
            nonVisualGroupShapeProperties6.Append(applicationNonVisualDrawingProperties29);

            GroupShapeProperties groupShapeProperties6 = new GroupShapeProperties();

            A.TransformGroup transformGroup6 = new A.TransformGroup();
            A.Offset offset20 = new A.Offset(){ X = 0L, Y = 0L };
            A.Extents extents20 = new A.Extents(){ Cx = 0L, Cy = 0L };
            A.ChildOffset childOffset6 = new A.ChildOffset(){ X = 0L, Y = 0L };
            A.ChildExtents childExtents6 = new A.ChildExtents(){ Cx = 0L, Cy = 0L };

            transformGroup6.Append(offset20);
            transformGroup6.Append(extents20);
            transformGroup6.Append(childOffset6);
            transformGroup6.Append(childExtents6);

            groupShapeProperties6.Append(transformGroup6);

            Shape shape22 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties22 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties30 = new NonVisualDrawingProperties(){ Id = (UInt32Value)2U, Name = "Date Placeholder 1" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties22 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks22 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties22.Append(shapeLocks22);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties30 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape22 = new PlaceholderShape(){ Type = PlaceholderValues.DateAndTime, Size = PlaceholderSizeValues.Half, Index = (UInt32Value)10U };

            applicationNonVisualDrawingProperties30.Append(placeholderShape22);

            nonVisualShapeProperties22.Append(nonVisualDrawingProperties30);
            nonVisualShapeProperties22.Append(nonVisualShapeDrawingProperties22);
            nonVisualShapeProperties22.Append(applicationNonVisualDrawingProperties30);
            ShapeProperties shapeProperties23 = new ShapeProperties();

            TextBody textBody22 = new TextBody();
            A.BodyProperties bodyProperties22 = new A.BodyProperties();
            A.ListStyle listStyle22 = new A.ListStyle();

            A.Paragraph paragraph30 = new A.Paragraph();

            A.Field field9 = new A.Field(){ Id = "{1D065981-0097-4BA9-A692-9FF652AFAF35}", Type = "datetimeFigureOut" };

            A.RunProperties runProperties26 = new A.RunProperties(){ Language = "cs-CZ" };
            runProperties26.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            A.Text text26 = new A.Text();
            text26.Text = "14.03.2023";

            field9.Append(runProperties26);
            field9.Append(text26);
            A.EndParagraphRunProperties endParagraphRunProperties20 = new A.EndParagraphRunProperties(){ Language = "cs-CZ" };

            paragraph30.Append(field9);
            paragraph30.Append(endParagraphRunProperties20);

            textBody22.Append(bodyProperties22);
            textBody22.Append(listStyle22);
            textBody22.Append(paragraph30);

            shape22.Append(nonVisualShapeProperties22);
            shape22.Append(shapeProperties23);
            shape22.Append(textBody22);

            Shape shape23 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties23 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties31 = new NonVisualDrawingProperties(){ Id = (UInt32Value)3U, Name = "Footer Placeholder 2" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties23 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks23 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties23.Append(shapeLocks23);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties31 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape23 = new PlaceholderShape(){ Type = PlaceholderValues.Footer, Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)11U };

            applicationNonVisualDrawingProperties31.Append(placeholderShape23);

            nonVisualShapeProperties23.Append(nonVisualDrawingProperties31);
            nonVisualShapeProperties23.Append(nonVisualShapeDrawingProperties23);
            nonVisualShapeProperties23.Append(applicationNonVisualDrawingProperties31);
            ShapeProperties shapeProperties24 = new ShapeProperties();

            TextBody textBody23 = new TextBody();
            A.BodyProperties bodyProperties23 = new A.BodyProperties();
            A.ListStyle listStyle23 = new A.ListStyle();

            A.Paragraph paragraph31 = new A.Paragraph();
            A.EndParagraphRunProperties endParagraphRunProperties21 = new A.EndParagraphRunProperties(){ Language = "cs-CZ" };

            paragraph31.Append(endParagraphRunProperties21);

            textBody23.Append(bodyProperties23);
            textBody23.Append(listStyle23);
            textBody23.Append(paragraph31);

            shape23.Append(nonVisualShapeProperties23);
            shape23.Append(shapeProperties24);
            shape23.Append(textBody23);

            Shape shape24 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties24 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties32 = new NonVisualDrawingProperties(){ Id = (UInt32Value)4U, Name = "Slide Number Placeholder 3" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties24 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks24 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties24.Append(shapeLocks24);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties32 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape24 = new PlaceholderShape(){ Type = PlaceholderValues.SlideNumber, Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)12U };

            applicationNonVisualDrawingProperties32.Append(placeholderShape24);

            nonVisualShapeProperties24.Append(nonVisualDrawingProperties32);
            nonVisualShapeProperties24.Append(nonVisualShapeDrawingProperties24);
            nonVisualShapeProperties24.Append(applicationNonVisualDrawingProperties32);
            ShapeProperties shapeProperties25 = new ShapeProperties();

            TextBody textBody24 = new TextBody();
            A.BodyProperties bodyProperties24 = new A.BodyProperties();
            A.ListStyle listStyle24 = new A.ListStyle();

            A.Paragraph paragraph32 = new A.Paragraph();

            A.Field field10 = new A.Field(){ Id = "{5D7F2B33-3E35-4F97-A448-9FEB151DF743}", Type = "slidenum" };

            A.RunProperties runProperties27 = new A.RunProperties(){ Language = "cs-CZ" };
            runProperties27.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            A.Text text27 = new A.Text();
            text27.Text = "‹#›";

            field10.Append(runProperties27);
            field10.Append(text27);
            A.EndParagraphRunProperties endParagraphRunProperties22 = new A.EndParagraphRunProperties(){ Language = "cs-CZ" };

            paragraph32.Append(field10);
            paragraph32.Append(endParagraphRunProperties22);

            textBody24.Append(bodyProperties24);
            textBody24.Append(listStyle24);
            textBody24.Append(paragraph32);

            shape24.Append(nonVisualShapeProperties24);
            shape24.Append(shapeProperties25);
            shape24.Append(textBody24);

            shapeTree6.Append(nonVisualGroupShapeProperties6);
            shapeTree6.Append(groupShapeProperties6);
            shapeTree6.Append(shape22);
            shapeTree6.Append(shape23);
            shapeTree6.Append(shape24);

            CommonSlideDataExtensionList commonSlideDataExtensionList6 = new CommonSlideDataExtensionList();

            CommonSlideDataExtension commonSlideDataExtension6 = new CommonSlideDataExtension(){ Uri = "{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}" };

            P14.CreationId creationId6 = new P14.CreationId(){ Val = (UInt32Value)3347405168U };
            creationId6.AddNamespaceDeclaration("p14", "http://schemas.microsoft.com/office/powerpoint/2010/main");

            commonSlideDataExtension6.Append(creationId6);

            commonSlideDataExtensionList6.Append(commonSlideDataExtension6);

            commonSlideData6.Append(shapeTree6);
            commonSlideData6.Append(commonSlideDataExtensionList6);

            ColorMapOverride colorMapOverride5 = new ColorMapOverride();
            A.MasterColorMapping masterColorMapping5 = new A.MasterColorMapping();

            colorMapOverride5.Append(masterColorMapping5);

            slideLayout4.Append(commonSlideData6);
            slideLayout4.Append(colorMapOverride5);

            slideLayoutPart4.SlideLayout = slideLayout4;
        }

        // Generates content of themePart1.
        private void GenerateThemePart1Content(ThemePart themePart1)
        {
            A.Theme theme1 = new A.Theme(){ Name = "Motiv Office" };
            theme1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

            A.ThemeElements themeElements1 = new A.ThemeElements();

            A.ColorScheme colorScheme1 = new A.ColorScheme(){ Name = "Motiv Office" };

            A.Dark1Color dark1Color1 = new A.Dark1Color();
            A.SystemColor systemColor1 = new A.SystemColor(){ Val = A.SystemColorValues.WindowText, LastColor = "000000" };

            dark1Color1.Append(systemColor1);

            A.Light1Color light1Color1 = new A.Light1Color();
            A.SystemColor systemColor2 = new A.SystemColor(){ Val = A.SystemColorValues.Window, LastColor = "FFFFFF" };

            light1Color1.Append(systemColor2);

            A.Dark2Color dark2Color1 = new A.Dark2Color();
            A.RgbColorModelHex rgbColorModelHex1 = new A.RgbColorModelHex(){ Val = "44546A" };

            dark2Color1.Append(rgbColorModelHex1);

            A.Light2Color light2Color1 = new A.Light2Color();
            A.RgbColorModelHex rgbColorModelHex2 = new A.RgbColorModelHex(){ Val = "E7E6E6" };

            light2Color1.Append(rgbColorModelHex2);

            A.Accent1Color accent1Color1 = new A.Accent1Color();
            A.RgbColorModelHex rgbColorModelHex3 = new A.RgbColorModelHex(){ Val = "4472C4" };

            accent1Color1.Append(rgbColorModelHex3);

            A.Accent2Color accent2Color1 = new A.Accent2Color();
            A.RgbColorModelHex rgbColorModelHex4 = new A.RgbColorModelHex(){ Val = "ED7D31" };

            accent2Color1.Append(rgbColorModelHex4);

            A.Accent3Color accent3Color1 = new A.Accent3Color();
            A.RgbColorModelHex rgbColorModelHex5 = new A.RgbColorModelHex(){ Val = "A5A5A5" };

            accent3Color1.Append(rgbColorModelHex5);

            A.Accent4Color accent4Color1 = new A.Accent4Color();
            A.RgbColorModelHex rgbColorModelHex6 = new A.RgbColorModelHex(){ Val = "FFC000" };

            accent4Color1.Append(rgbColorModelHex6);

            A.Accent5Color accent5Color1 = new A.Accent5Color();
            A.RgbColorModelHex rgbColorModelHex7 = new A.RgbColorModelHex(){ Val = "5B9BD5" };

            accent5Color1.Append(rgbColorModelHex7);

            A.Accent6Color accent6Color1 = new A.Accent6Color();
            A.RgbColorModelHex rgbColorModelHex8 = new A.RgbColorModelHex(){ Val = "70AD47" };

            accent6Color1.Append(rgbColorModelHex8);

            A.Hyperlink hyperlink1 = new A.Hyperlink();
            A.RgbColorModelHex rgbColorModelHex9 = new A.RgbColorModelHex(){ Val = "0563C1" };

            hyperlink1.Append(rgbColorModelHex9);

            A.FollowedHyperlinkColor followedHyperlinkColor1 = new A.FollowedHyperlinkColor();
            A.RgbColorModelHex rgbColorModelHex10 = new A.RgbColorModelHex(){ Val = "954F72" };

            followedHyperlinkColor1.Append(rgbColorModelHex10);

            colorScheme1.Append(dark1Color1);
            colorScheme1.Append(light1Color1);
            colorScheme1.Append(dark2Color1);
            colorScheme1.Append(light2Color1);
            colorScheme1.Append(accent1Color1);
            colorScheme1.Append(accent2Color1);
            colorScheme1.Append(accent3Color1);
            colorScheme1.Append(accent4Color1);
            colorScheme1.Append(accent5Color1);
            colorScheme1.Append(accent6Color1);
            colorScheme1.Append(hyperlink1);
            colorScheme1.Append(followedHyperlinkColor1);

            A.FontScheme fontScheme1 = new A.FontScheme(){ Name = "Motiv Office" };

            A.MajorFont majorFont1 = new A.MajorFont();
            A.LatinFont latinFont29 = new A.LatinFont(){ Typeface = "Calibri Light", Panose = "020F0302020204030204" };
            A.EastAsianFont eastAsianFont29 = new A.EastAsianFont(){ Typeface = "" };
            A.ComplexScriptFont complexScriptFont29 = new A.ComplexScriptFont(){ Typeface = "" };
            A.SupplementalFont supplementalFont1 = new A.SupplementalFont(){ Script = "Jpan", Typeface = "游ゴシック Light" };
            A.SupplementalFont supplementalFont2 = new A.SupplementalFont(){ Script = "Hang", Typeface = "맑은 고딕" };
            A.SupplementalFont supplementalFont3 = new A.SupplementalFont(){ Script = "Hans", Typeface = "等线 Light" };
            A.SupplementalFont supplementalFont4 = new A.SupplementalFont(){ Script = "Hant", Typeface = "新細明體" };
            A.SupplementalFont supplementalFont5 = new A.SupplementalFont(){ Script = "Arab", Typeface = "Times New Roman" };
            A.SupplementalFont supplementalFont6 = new A.SupplementalFont(){ Script = "Hebr", Typeface = "Times New Roman" };
            A.SupplementalFont supplementalFont7 = new A.SupplementalFont(){ Script = "Thai", Typeface = "Angsana New" };
            A.SupplementalFont supplementalFont8 = new A.SupplementalFont(){ Script = "Ethi", Typeface = "Nyala" };
            A.SupplementalFont supplementalFont9 = new A.SupplementalFont(){ Script = "Beng", Typeface = "Vrinda" };
            A.SupplementalFont supplementalFont10 = new A.SupplementalFont(){ Script = "Gujr", Typeface = "Shruti" };
            A.SupplementalFont supplementalFont11 = new A.SupplementalFont(){ Script = "Khmr", Typeface = "MoolBoran" };
            A.SupplementalFont supplementalFont12 = new A.SupplementalFont(){ Script = "Knda", Typeface = "Tunga" };
            A.SupplementalFont supplementalFont13 = new A.SupplementalFont(){ Script = "Guru", Typeface = "Raavi" };
            A.SupplementalFont supplementalFont14 = new A.SupplementalFont(){ Script = "Cans", Typeface = "Euphemia" };
            A.SupplementalFont supplementalFont15 = new A.SupplementalFont(){ Script = "Cher", Typeface = "Plantagenet Cherokee" };
            A.SupplementalFont supplementalFont16 = new A.SupplementalFont(){ Script = "Yiii", Typeface = "Microsoft Yi Baiti" };
            A.SupplementalFont supplementalFont17 = new A.SupplementalFont(){ Script = "Tibt", Typeface = "Microsoft Himalaya" };
            A.SupplementalFont supplementalFont18 = new A.SupplementalFont(){ Script = "Thaa", Typeface = "MV Boli" };
            A.SupplementalFont supplementalFont19 = new A.SupplementalFont(){ Script = "Deva", Typeface = "Mangal" };
            A.SupplementalFont supplementalFont20 = new A.SupplementalFont(){ Script = "Telu", Typeface = "Gautami" };
            A.SupplementalFont supplementalFont21 = new A.SupplementalFont(){ Script = "Taml", Typeface = "Latha" };
            A.SupplementalFont supplementalFont22 = new A.SupplementalFont(){ Script = "Syrc", Typeface = "Estrangelo Edessa" };
            A.SupplementalFont supplementalFont23 = new A.SupplementalFont(){ Script = "Orya", Typeface = "Kalinga" };
            A.SupplementalFont supplementalFont24 = new A.SupplementalFont(){ Script = "Mlym", Typeface = "Kartika" };
            A.SupplementalFont supplementalFont25 = new A.SupplementalFont(){ Script = "Laoo", Typeface = "DokChampa" };
            A.SupplementalFont supplementalFont26 = new A.SupplementalFont(){ Script = "Sinh", Typeface = "Iskoola Pota" };
            A.SupplementalFont supplementalFont27 = new A.SupplementalFont(){ Script = "Mong", Typeface = "Mongolian Baiti" };
            A.SupplementalFont supplementalFont28 = new A.SupplementalFont(){ Script = "Viet", Typeface = "Times New Roman" };
            A.SupplementalFont supplementalFont29 = new A.SupplementalFont(){ Script = "Uigh", Typeface = "Microsoft Uighur" };
            A.SupplementalFont supplementalFont30 = new A.SupplementalFont(){ Script = "Geor", Typeface = "Sylfaen" };

            majorFont1.Append(latinFont29);
            majorFont1.Append(eastAsianFont29);
            majorFont1.Append(complexScriptFont29);
            majorFont1.Append(supplementalFont1);
            majorFont1.Append(supplementalFont2);
            majorFont1.Append(supplementalFont3);
            majorFont1.Append(supplementalFont4);
            majorFont1.Append(supplementalFont5);
            majorFont1.Append(supplementalFont6);
            majorFont1.Append(supplementalFont7);
            majorFont1.Append(supplementalFont8);
            majorFont1.Append(supplementalFont9);
            majorFont1.Append(supplementalFont10);
            majorFont1.Append(supplementalFont11);
            majorFont1.Append(supplementalFont12);
            majorFont1.Append(supplementalFont13);
            majorFont1.Append(supplementalFont14);
            majorFont1.Append(supplementalFont15);
            majorFont1.Append(supplementalFont16);
            majorFont1.Append(supplementalFont17);
            majorFont1.Append(supplementalFont18);
            majorFont1.Append(supplementalFont19);
            majorFont1.Append(supplementalFont20);
            majorFont1.Append(supplementalFont21);
            majorFont1.Append(supplementalFont22);
            majorFont1.Append(supplementalFont23);
            majorFont1.Append(supplementalFont24);
            majorFont1.Append(supplementalFont25);
            majorFont1.Append(supplementalFont26);
            majorFont1.Append(supplementalFont27);
            majorFont1.Append(supplementalFont28);
            majorFont1.Append(supplementalFont29);
            majorFont1.Append(supplementalFont30);

            A.MinorFont minorFont1 = new A.MinorFont();
            A.LatinFont latinFont30 = new A.LatinFont(){ Typeface = "Calibri", Panose = "020F0502020204030204" };
            A.EastAsianFont eastAsianFont30 = new A.EastAsianFont(){ Typeface = "" };
            A.ComplexScriptFont complexScriptFont30 = new A.ComplexScriptFont(){ Typeface = "" };
            A.SupplementalFont supplementalFont31 = new A.SupplementalFont(){ Script = "Jpan", Typeface = "游ゴシック" };
            A.SupplementalFont supplementalFont32 = new A.SupplementalFont(){ Script = "Hang", Typeface = "맑은 고딕" };
            A.SupplementalFont supplementalFont33 = new A.SupplementalFont(){ Script = "Hans", Typeface = "等线" };
            A.SupplementalFont supplementalFont34 = new A.SupplementalFont(){ Script = "Hant", Typeface = "新細明體" };
            A.SupplementalFont supplementalFont35 = new A.SupplementalFont(){ Script = "Arab", Typeface = "Arial" };
            A.SupplementalFont supplementalFont36 = new A.SupplementalFont(){ Script = "Hebr", Typeface = "Arial" };
            A.SupplementalFont supplementalFont37 = new A.SupplementalFont(){ Script = "Thai", Typeface = "Cordia New" };
            A.SupplementalFont supplementalFont38 = new A.SupplementalFont(){ Script = "Ethi", Typeface = "Nyala" };
            A.SupplementalFont supplementalFont39 = new A.SupplementalFont(){ Script = "Beng", Typeface = "Vrinda" };
            A.SupplementalFont supplementalFont40 = new A.SupplementalFont(){ Script = "Gujr", Typeface = "Shruti" };
            A.SupplementalFont supplementalFont41 = new A.SupplementalFont(){ Script = "Khmr", Typeface = "DaunPenh" };
            A.SupplementalFont supplementalFont42 = new A.SupplementalFont(){ Script = "Knda", Typeface = "Tunga" };
            A.SupplementalFont supplementalFont43 = new A.SupplementalFont(){ Script = "Guru", Typeface = "Raavi" };
            A.SupplementalFont supplementalFont44 = new A.SupplementalFont(){ Script = "Cans", Typeface = "Euphemia" };
            A.SupplementalFont supplementalFont45 = new A.SupplementalFont(){ Script = "Cher", Typeface = "Plantagenet Cherokee" };
            A.SupplementalFont supplementalFont46 = new A.SupplementalFont(){ Script = "Yiii", Typeface = "Microsoft Yi Baiti" };
            A.SupplementalFont supplementalFont47 = new A.SupplementalFont(){ Script = "Tibt", Typeface = "Microsoft Himalaya" };
            A.SupplementalFont supplementalFont48 = new A.SupplementalFont(){ Script = "Thaa", Typeface = "MV Boli" };
            A.SupplementalFont supplementalFont49 = new A.SupplementalFont(){ Script = "Deva", Typeface = "Mangal" };
            A.SupplementalFont supplementalFont50 = new A.SupplementalFont(){ Script = "Telu", Typeface = "Gautami" };
            A.SupplementalFont supplementalFont51 = new A.SupplementalFont(){ Script = "Taml", Typeface = "Latha" };
            A.SupplementalFont supplementalFont52 = new A.SupplementalFont(){ Script = "Syrc", Typeface = "Estrangelo Edessa" };
            A.SupplementalFont supplementalFont53 = new A.SupplementalFont(){ Script = "Orya", Typeface = "Kalinga" };
            A.SupplementalFont supplementalFont54 = new A.SupplementalFont(){ Script = "Mlym", Typeface = "Kartika" };
            A.SupplementalFont supplementalFont55 = new A.SupplementalFont(){ Script = "Laoo", Typeface = "DokChampa" };
            A.SupplementalFont supplementalFont56 = new A.SupplementalFont(){ Script = "Sinh", Typeface = "Iskoola Pota" };
            A.SupplementalFont supplementalFont57 = new A.SupplementalFont(){ Script = "Mong", Typeface = "Mongolian Baiti" };
            A.SupplementalFont supplementalFont58 = new A.SupplementalFont(){ Script = "Viet", Typeface = "Arial" };
            A.SupplementalFont supplementalFont59 = new A.SupplementalFont(){ Script = "Uigh", Typeface = "Microsoft Uighur" };
            A.SupplementalFont supplementalFont60 = new A.SupplementalFont(){ Script = "Geor", Typeface = "Sylfaen" };

            minorFont1.Append(latinFont30);
            minorFont1.Append(eastAsianFont30);
            minorFont1.Append(complexScriptFont30);
            minorFont1.Append(supplementalFont31);
            minorFont1.Append(supplementalFont32);
            minorFont1.Append(supplementalFont33);
            minorFont1.Append(supplementalFont34);
            minorFont1.Append(supplementalFont35);
            minorFont1.Append(supplementalFont36);
            minorFont1.Append(supplementalFont37);
            minorFont1.Append(supplementalFont38);
            minorFont1.Append(supplementalFont39);
            minorFont1.Append(supplementalFont40);
            minorFont1.Append(supplementalFont41);
            minorFont1.Append(supplementalFont42);
            minorFont1.Append(supplementalFont43);
            minorFont1.Append(supplementalFont44);
            minorFont1.Append(supplementalFont45);
            minorFont1.Append(supplementalFont46);
            minorFont1.Append(supplementalFont47);
            minorFont1.Append(supplementalFont48);
            minorFont1.Append(supplementalFont49);
            minorFont1.Append(supplementalFont50);
            minorFont1.Append(supplementalFont51);
            minorFont1.Append(supplementalFont52);
            minorFont1.Append(supplementalFont53);
            minorFont1.Append(supplementalFont54);
            minorFont1.Append(supplementalFont55);
            minorFont1.Append(supplementalFont56);
            minorFont1.Append(supplementalFont57);
            minorFont1.Append(supplementalFont58);
            minorFont1.Append(supplementalFont59);
            minorFont1.Append(supplementalFont60);

            fontScheme1.Append(majorFont1);
            fontScheme1.Append(minorFont1);

            A.FormatScheme formatScheme1 = new A.FormatScheme(){ Name = "Motiv Office" };

            A.FillStyleList fillStyleList1 = new A.FillStyleList();

            A.SolidFill solidFill41 = new A.SolidFill();
            A.SchemeColor schemeColor42 = new A.SchemeColor(){ Val = A.SchemeColorValues.PhColor };

            solidFill41.Append(schemeColor42);

            A.GradientFill gradientFill1 = new A.GradientFill(){ RotateWithShape = true };

            A.GradientStopList gradientStopList1 = new A.GradientStopList();

            A.GradientStop gradientStop1 = new A.GradientStop(){ Position = 0 };

            A.SchemeColor schemeColor43 = new A.SchemeColor(){ Val = A.SchemeColorValues.PhColor };
            A.LuminanceModulation luminanceModulation1 = new A.LuminanceModulation(){ Val = 110000 };
            A.SaturationModulation saturationModulation1 = new A.SaturationModulation(){ Val = 105000 };
            A.Tint tint12 = new A.Tint(){ Val = 67000 };

            schemeColor43.Append(luminanceModulation1);
            schemeColor43.Append(saturationModulation1);
            schemeColor43.Append(tint12);

            gradientStop1.Append(schemeColor43);

            A.GradientStop gradientStop2 = new A.GradientStop(){ Position = 50000 };

            A.SchemeColor schemeColor44 = new A.SchemeColor(){ Val = A.SchemeColorValues.PhColor };
            A.LuminanceModulation luminanceModulation2 = new A.LuminanceModulation(){ Val = 105000 };
            A.SaturationModulation saturationModulation2 = new A.SaturationModulation(){ Val = 103000 };
            A.Tint tint13 = new A.Tint(){ Val = 73000 };

            schemeColor44.Append(luminanceModulation2);
            schemeColor44.Append(saturationModulation2);
            schemeColor44.Append(tint13);

            gradientStop2.Append(schemeColor44);

            A.GradientStop gradientStop3 = new A.GradientStop(){ Position = 100000 };

            A.SchemeColor schemeColor45 = new A.SchemeColor(){ Val = A.SchemeColorValues.PhColor };
            A.LuminanceModulation luminanceModulation3 = new A.LuminanceModulation(){ Val = 105000 };
            A.SaturationModulation saturationModulation3 = new A.SaturationModulation(){ Val = 109000 };
            A.Tint tint14 = new A.Tint(){ Val = 81000 };

            schemeColor45.Append(luminanceModulation3);
            schemeColor45.Append(saturationModulation3);
            schemeColor45.Append(tint14);

            gradientStop3.Append(schemeColor45);

            gradientStopList1.Append(gradientStop1);
            gradientStopList1.Append(gradientStop2);
            gradientStopList1.Append(gradientStop3);
            A.LinearGradientFill linearGradientFill1 = new A.LinearGradientFill(){ Angle = 5400000, Scaled = false };

            gradientFill1.Append(gradientStopList1);
            gradientFill1.Append(linearGradientFill1);

            A.GradientFill gradientFill2 = new A.GradientFill(){ RotateWithShape = true };

            A.GradientStopList gradientStopList2 = new A.GradientStopList();

            A.GradientStop gradientStop4 = new A.GradientStop(){ Position = 0 };

            A.SchemeColor schemeColor46 = new A.SchemeColor(){ Val = A.SchemeColorValues.PhColor };
            A.SaturationModulation saturationModulation4 = new A.SaturationModulation(){ Val = 103000 };
            A.LuminanceModulation luminanceModulation4 = new A.LuminanceModulation(){ Val = 102000 };
            A.Tint tint15 = new A.Tint(){ Val = 94000 };

            schemeColor46.Append(saturationModulation4);
            schemeColor46.Append(luminanceModulation4);
            schemeColor46.Append(tint15);

            gradientStop4.Append(schemeColor46);

            A.GradientStop gradientStop5 = new A.GradientStop(){ Position = 50000 };

            A.SchemeColor schemeColor47 = new A.SchemeColor(){ Val = A.SchemeColorValues.PhColor };
            A.SaturationModulation saturationModulation5 = new A.SaturationModulation(){ Val = 110000 };
            A.LuminanceModulation luminanceModulation5 = new A.LuminanceModulation(){ Val = 100000 };
            A.Shade shade1 = new A.Shade(){ Val = 100000 };

            schemeColor47.Append(saturationModulation5);
            schemeColor47.Append(luminanceModulation5);
            schemeColor47.Append(shade1);

            gradientStop5.Append(schemeColor47);

            A.GradientStop gradientStop6 = new A.GradientStop(){ Position = 100000 };

            A.SchemeColor schemeColor48 = new A.SchemeColor(){ Val = A.SchemeColorValues.PhColor };
            A.LuminanceModulation luminanceModulation6 = new A.LuminanceModulation(){ Val = 99000 };
            A.SaturationModulation saturationModulation6 = new A.SaturationModulation(){ Val = 120000 };
            A.Shade shade2 = new A.Shade(){ Val = 78000 };

            schemeColor48.Append(luminanceModulation6);
            schemeColor48.Append(saturationModulation6);
            schemeColor48.Append(shade2);

            gradientStop6.Append(schemeColor48);

            gradientStopList2.Append(gradientStop4);
            gradientStopList2.Append(gradientStop5);
            gradientStopList2.Append(gradientStop6);
            A.LinearGradientFill linearGradientFill2 = new A.LinearGradientFill(){ Angle = 5400000, Scaled = false };

            gradientFill2.Append(gradientStopList2);
            gradientFill2.Append(linearGradientFill2);

            fillStyleList1.Append(solidFill41);
            fillStyleList1.Append(gradientFill1);
            fillStyleList1.Append(gradientFill2);

            A.LineStyleList lineStyleList1 = new A.LineStyleList();

            A.Outline outline1 = new A.Outline(){ Width = 6350, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill42 = new A.SolidFill();
            A.SchemeColor schemeColor49 = new A.SchemeColor(){ Val = A.SchemeColorValues.PhColor };

            solidFill42.Append(schemeColor49);
            A.PresetDash presetDash1 = new A.PresetDash(){ Val = A.PresetLineDashValues.Solid };
            A.Miter miter1 = new A.Miter(){ Limit = 800000 };

            outline1.Append(solidFill42);
            outline1.Append(presetDash1);
            outline1.Append(miter1);

            A.Outline outline2 = new A.Outline(){ Width = 12700, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill43 = new A.SolidFill();
            A.SchemeColor schemeColor50 = new A.SchemeColor(){ Val = A.SchemeColorValues.PhColor };

            solidFill43.Append(schemeColor50);
            A.PresetDash presetDash2 = new A.PresetDash(){ Val = A.PresetLineDashValues.Solid };
            A.Miter miter2 = new A.Miter(){ Limit = 800000 };

            outline2.Append(solidFill43);
            outline2.Append(presetDash2);
            outline2.Append(miter2);

            A.Outline outline3 = new A.Outline(){ Width = 19050, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill44 = new A.SolidFill();
            A.SchemeColor schemeColor51 = new A.SchemeColor(){ Val = A.SchemeColorValues.PhColor };

            solidFill44.Append(schemeColor51);
            A.PresetDash presetDash3 = new A.PresetDash(){ Val = A.PresetLineDashValues.Solid };
            A.Miter miter3 = new A.Miter(){ Limit = 800000 };

            outline3.Append(solidFill44);
            outline3.Append(presetDash3);
            outline3.Append(miter3);

            lineStyleList1.Append(outline1);
            lineStyleList1.Append(outline2);
            lineStyleList1.Append(outline3);

            A.EffectStyleList effectStyleList1 = new A.EffectStyleList();

            A.EffectStyle effectStyle1 = new A.EffectStyle();
            A.EffectList effectList1 = new A.EffectList();

            effectStyle1.Append(effectList1);

            A.EffectStyle effectStyle2 = new A.EffectStyle();
            A.EffectList effectList2 = new A.EffectList();

            effectStyle2.Append(effectList2);

            A.EffectStyle effectStyle3 = new A.EffectStyle();

            A.EffectList effectList3 = new A.EffectList();

            A.OuterShadow outerShadow1 = new A.OuterShadow(){ BlurRadius = 57150L, Distance = 19050L, Direction = 5400000, Alignment = A.RectangleAlignmentValues.Center, RotateWithShape = false };

            A.RgbColorModelHex rgbColorModelHex11 = new A.RgbColorModelHex(){ Val = "000000" };
            A.Alpha alpha1 = new A.Alpha(){ Val = 63000 };

            rgbColorModelHex11.Append(alpha1);

            outerShadow1.Append(rgbColorModelHex11);

            effectList3.Append(outerShadow1);

            effectStyle3.Append(effectList3);

            effectStyleList1.Append(effectStyle1);
            effectStyleList1.Append(effectStyle2);
            effectStyleList1.Append(effectStyle3);

            A.BackgroundFillStyleList backgroundFillStyleList1 = new A.BackgroundFillStyleList();

            A.SolidFill solidFill45 = new A.SolidFill();
            A.SchemeColor schemeColor52 = new A.SchemeColor(){ Val = A.SchemeColorValues.PhColor };

            solidFill45.Append(schemeColor52);

            A.SolidFill solidFill46 = new A.SolidFill();

            A.SchemeColor schemeColor53 = new A.SchemeColor(){ Val = A.SchemeColorValues.PhColor };
            A.Tint tint16 = new A.Tint(){ Val = 95000 };
            A.SaturationModulation saturationModulation7 = new A.SaturationModulation(){ Val = 170000 };

            schemeColor53.Append(tint16);
            schemeColor53.Append(saturationModulation7);

            solidFill46.Append(schemeColor53);

            A.GradientFill gradientFill3 = new A.GradientFill(){ RotateWithShape = true };

            A.GradientStopList gradientStopList3 = new A.GradientStopList();

            A.GradientStop gradientStop7 = new A.GradientStop(){ Position = 0 };

            A.SchemeColor schemeColor54 = new A.SchemeColor(){ Val = A.SchemeColorValues.PhColor };
            A.Tint tint17 = new A.Tint(){ Val = 93000 };
            A.SaturationModulation saturationModulation8 = new A.SaturationModulation(){ Val = 150000 };
            A.Shade shade3 = new A.Shade(){ Val = 98000 };
            A.LuminanceModulation luminanceModulation7 = new A.LuminanceModulation(){ Val = 102000 };

            schemeColor54.Append(tint17);
            schemeColor54.Append(saturationModulation8);
            schemeColor54.Append(shade3);
            schemeColor54.Append(luminanceModulation7);

            gradientStop7.Append(schemeColor54);

            A.GradientStop gradientStop8 = new A.GradientStop(){ Position = 50000 };

            A.SchemeColor schemeColor55 = new A.SchemeColor(){ Val = A.SchemeColorValues.PhColor };
            A.Tint tint18 = new A.Tint(){ Val = 98000 };
            A.SaturationModulation saturationModulation9 = new A.SaturationModulation(){ Val = 130000 };
            A.Shade shade4 = new A.Shade(){ Val = 90000 };
            A.LuminanceModulation luminanceModulation8 = new A.LuminanceModulation(){ Val = 103000 };

            schemeColor55.Append(tint18);
            schemeColor55.Append(saturationModulation9);
            schemeColor55.Append(shade4);
            schemeColor55.Append(luminanceModulation8);

            gradientStop8.Append(schemeColor55);

            A.GradientStop gradientStop9 = new A.GradientStop(){ Position = 100000 };

            A.SchemeColor schemeColor56 = new A.SchemeColor(){ Val = A.SchemeColorValues.PhColor };
            A.Shade shade5 = new A.Shade(){ Val = 63000 };
            A.SaturationModulation saturationModulation10 = new A.SaturationModulation(){ Val = 120000 };

            schemeColor56.Append(shade5);
            schemeColor56.Append(saturationModulation10);

            gradientStop9.Append(schemeColor56);

            gradientStopList3.Append(gradientStop7);
            gradientStopList3.Append(gradientStop8);
            gradientStopList3.Append(gradientStop9);
            A.LinearGradientFill linearGradientFill3 = new A.LinearGradientFill(){ Angle = 5400000, Scaled = false };

            gradientFill3.Append(gradientStopList3);
            gradientFill3.Append(linearGradientFill3);

            backgroundFillStyleList1.Append(solidFill45);
            backgroundFillStyleList1.Append(solidFill46);
            backgroundFillStyleList1.Append(gradientFill3);

            formatScheme1.Append(fillStyleList1);
            formatScheme1.Append(lineStyleList1);
            formatScheme1.Append(effectStyleList1);
            formatScheme1.Append(backgroundFillStyleList1);

            themeElements1.Append(colorScheme1);
            themeElements1.Append(fontScheme1);
            themeElements1.Append(formatScheme1);
            A.ObjectDefaults objectDefaults1 = new A.ObjectDefaults();
            A.ExtraColorSchemeList extraColorSchemeList1 = new A.ExtraColorSchemeList();

            A.OfficeStyleSheetExtensionList officeStyleSheetExtensionList1 = new A.OfficeStyleSheetExtensionList();

            A.OfficeStyleSheetExtension officeStyleSheetExtension1 = new A.OfficeStyleSheetExtension(){ Uri = "{05A4C25C-085E-4340-85A3-A5531E510DB2}" };

            Thm15.ThemeFamily themeFamily1 = new Thm15.ThemeFamily(){ Name = "Office Theme 2013 - 2022", Id = "{62F939B6-93AF-4DB8-9C6B-D6C7DFDC589F}", Vid = "{4A3C46E8-61CC-4603-A589-7422A47A8E4A}" };
            themeFamily1.AddNamespaceDeclaration("thm15", "http://schemas.microsoft.com/office/thememl/2012/main");

            officeStyleSheetExtension1.Append(themeFamily1);

            officeStyleSheetExtensionList1.Append(officeStyleSheetExtension1);

            theme1.Append(themeElements1);
            theme1.Append(objectDefaults1);
            theme1.Append(extraColorSchemeList1);
            theme1.Append(officeStyleSheetExtensionList1);

            themePart1.Theme = theme1;
        }

        // Generates content of slideLayoutPart5.
        private void GenerateSlideLayoutPart5Content(SlideLayoutPart slideLayoutPart5)
        {
            SlideLayout slideLayout5 = new SlideLayout(){ Type = SlideLayoutValues.Object, Preserve = true };
            slideLayout5.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            slideLayout5.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            slideLayout5.AddNamespaceDeclaration("p", "http://schemas.openxmlformats.org/presentationml/2006/main");

            CommonSlideData commonSlideData7 = new CommonSlideData(){ Name = "Title and Content" };

            ShapeTree shapeTree7 = new ShapeTree();

            NonVisualGroupShapeProperties nonVisualGroupShapeProperties7 = new NonVisualGroupShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties33 = new NonVisualDrawingProperties(){ Id = (UInt32Value)1U, Name = "" };
            NonVisualGroupShapeDrawingProperties nonVisualGroupShapeDrawingProperties7 = new NonVisualGroupShapeDrawingProperties();
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties33 = new ApplicationNonVisualDrawingProperties();

            nonVisualGroupShapeProperties7.Append(nonVisualDrawingProperties33);
            nonVisualGroupShapeProperties7.Append(nonVisualGroupShapeDrawingProperties7);
            nonVisualGroupShapeProperties7.Append(applicationNonVisualDrawingProperties33);

            GroupShapeProperties groupShapeProperties7 = new GroupShapeProperties();

            A.TransformGroup transformGroup7 = new A.TransformGroup();
            A.Offset offset21 = new A.Offset(){ X = 0L, Y = 0L };
            A.Extents extents21 = new A.Extents(){ Cx = 0L, Cy = 0L };
            A.ChildOffset childOffset7 = new A.ChildOffset(){ X = 0L, Y = 0L };
            A.ChildExtents childExtents7 = new A.ChildExtents(){ Cx = 0L, Cy = 0L };

            transformGroup7.Append(offset21);
            transformGroup7.Append(extents21);
            transformGroup7.Append(childOffset7);
            transformGroup7.Append(childExtents7);

            groupShapeProperties7.Append(transformGroup7);

            Shape shape25 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties25 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties34 = new NonVisualDrawingProperties(){ Id = (UInt32Value)2U, Name = "Title 1" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties25 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks25 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties25.Append(shapeLocks25);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties34 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape25 = new PlaceholderShape(){ Type = PlaceholderValues.Title };

            applicationNonVisualDrawingProperties34.Append(placeholderShape25);

            nonVisualShapeProperties25.Append(nonVisualDrawingProperties34);
            nonVisualShapeProperties25.Append(nonVisualShapeDrawingProperties25);
            nonVisualShapeProperties25.Append(applicationNonVisualDrawingProperties34);
            ShapeProperties shapeProperties26 = new ShapeProperties();

            TextBody textBody25 = new TextBody();
            A.BodyProperties bodyProperties25 = new A.BodyProperties();
            A.ListStyle listStyle25 = new A.ListStyle();

            A.Paragraph paragraph33 = new A.Paragraph();

            A.Run run18 = new A.Run();
            A.RunProperties runProperties28 = new A.RunProperties(){ Language = "cs-CZ" };
            A.Text text28 = new A.Text();
            text28.Text = "Kliknutím lze upravit styl.";

            run18.Append(runProperties28);
            run18.Append(text28);
            A.EndParagraphRunProperties endParagraphRunProperties23 = new A.EndParagraphRunProperties(){ Language = "en-US", Dirty = false };

            paragraph33.Append(run18);
            paragraph33.Append(endParagraphRunProperties23);

            textBody25.Append(bodyProperties25);
            textBody25.Append(listStyle25);
            textBody25.Append(paragraph33);

            shape25.Append(nonVisualShapeProperties25);
            shape25.Append(shapeProperties26);
            shape25.Append(textBody25);

            Shape shape26 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties26 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties35 = new NonVisualDrawingProperties(){ Id = (UInt32Value)3U, Name = "Content Placeholder 2" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties26 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks26 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties26.Append(shapeLocks26);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties35 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape26 = new PlaceholderShape(){ Index = (UInt32Value)1U };

            applicationNonVisualDrawingProperties35.Append(placeholderShape26);

            nonVisualShapeProperties26.Append(nonVisualDrawingProperties35);
            nonVisualShapeProperties26.Append(nonVisualShapeDrawingProperties26);
            nonVisualShapeProperties26.Append(applicationNonVisualDrawingProperties35);
            ShapeProperties shapeProperties27 = new ShapeProperties();

            TextBody textBody26 = new TextBody();
            A.BodyProperties bodyProperties26 = new A.BodyProperties();
            A.ListStyle listStyle26 = new A.ListStyle();

            A.Paragraph paragraph34 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties13 = new A.ParagraphProperties(){ Level = 0 };

            A.Run run19 = new A.Run();
            A.RunProperties runProperties29 = new A.RunProperties(){ Language = "cs-CZ" };
            A.Text text29 = new A.Text();
            text29.Text = "Po kliknutí můžete upravovat styly textu v předloze.";

            run19.Append(runProperties29);
            run19.Append(text29);

            paragraph34.Append(paragraphProperties13);
            paragraph34.Append(run19);

            A.Paragraph paragraph35 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties14 = new A.ParagraphProperties(){ Level = 1 };

            A.Run run20 = new A.Run();
            A.RunProperties runProperties30 = new A.RunProperties(){ Language = "cs-CZ" };
            A.Text text30 = new A.Text();
            text30.Text = "Druhá úroveň";

            run20.Append(runProperties30);
            run20.Append(text30);

            paragraph35.Append(paragraphProperties14);
            paragraph35.Append(run20);

            A.Paragraph paragraph36 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties15 = new A.ParagraphProperties(){ Level = 2 };

            A.Run run21 = new A.Run();
            A.RunProperties runProperties31 = new A.RunProperties(){ Language = "cs-CZ" };
            A.Text text31 = new A.Text();
            text31.Text = "Třetí úroveň";

            run21.Append(runProperties31);
            run21.Append(text31);

            paragraph36.Append(paragraphProperties15);
            paragraph36.Append(run21);

            A.Paragraph paragraph37 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties16 = new A.ParagraphProperties(){ Level = 3 };

            A.Run run22 = new A.Run();
            A.RunProperties runProperties32 = new A.RunProperties(){ Language = "cs-CZ" };
            A.Text text32 = new A.Text();
            text32.Text = "Čtvrtá úroveň";

            run22.Append(runProperties32);
            run22.Append(text32);

            paragraph37.Append(paragraphProperties16);
            paragraph37.Append(run22);

            A.Paragraph paragraph38 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties17 = new A.ParagraphProperties(){ Level = 4 };

            A.Run run23 = new A.Run();
            A.RunProperties runProperties33 = new A.RunProperties(){ Language = "cs-CZ" };
            A.Text text33 = new A.Text();
            text33.Text = "Pátá úroveň";

            run23.Append(runProperties33);
            run23.Append(text33);
            A.EndParagraphRunProperties endParagraphRunProperties24 = new A.EndParagraphRunProperties(){ Language = "en-US", Dirty = false };

            paragraph38.Append(paragraphProperties17);
            paragraph38.Append(run23);
            paragraph38.Append(endParagraphRunProperties24);

            textBody26.Append(bodyProperties26);
            textBody26.Append(listStyle26);
            textBody26.Append(paragraph34);
            textBody26.Append(paragraph35);
            textBody26.Append(paragraph36);
            textBody26.Append(paragraph37);
            textBody26.Append(paragraph38);

            shape26.Append(nonVisualShapeProperties26);
            shape26.Append(shapeProperties27);
            shape26.Append(textBody26);

            Shape shape27 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties27 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties36 = new NonVisualDrawingProperties(){ Id = (UInt32Value)4U, Name = "Date Placeholder 3" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties27 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks27 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties27.Append(shapeLocks27);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties36 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape27 = new PlaceholderShape(){ Type = PlaceholderValues.DateAndTime, Size = PlaceholderSizeValues.Half, Index = (UInt32Value)10U };

            applicationNonVisualDrawingProperties36.Append(placeholderShape27);

            nonVisualShapeProperties27.Append(nonVisualDrawingProperties36);
            nonVisualShapeProperties27.Append(nonVisualShapeDrawingProperties27);
            nonVisualShapeProperties27.Append(applicationNonVisualDrawingProperties36);
            ShapeProperties shapeProperties28 = new ShapeProperties();

            TextBody textBody27 = new TextBody();
            A.BodyProperties bodyProperties27 = new A.BodyProperties();
            A.ListStyle listStyle27 = new A.ListStyle();

            A.Paragraph paragraph39 = new A.Paragraph();

            A.Field field11 = new A.Field(){ Id = "{1D065981-0097-4BA9-A692-9FF652AFAF35}", Type = "datetimeFigureOut" };

            A.RunProperties runProperties34 = new A.RunProperties(){ Language = "cs-CZ" };
            runProperties34.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            A.Text text34 = new A.Text();
            text34.Text = "14.03.2023";

            field11.Append(runProperties34);
            field11.Append(text34);
            A.EndParagraphRunProperties endParagraphRunProperties25 = new A.EndParagraphRunProperties(){ Language = "cs-CZ" };

            paragraph39.Append(field11);
            paragraph39.Append(endParagraphRunProperties25);

            textBody27.Append(bodyProperties27);
            textBody27.Append(listStyle27);
            textBody27.Append(paragraph39);

            shape27.Append(nonVisualShapeProperties27);
            shape27.Append(shapeProperties28);
            shape27.Append(textBody27);

            Shape shape28 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties28 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties37 = new NonVisualDrawingProperties(){ Id = (UInt32Value)5U, Name = "Footer Placeholder 4" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties28 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks28 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties28.Append(shapeLocks28);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties37 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape28 = new PlaceholderShape(){ Type = PlaceholderValues.Footer, Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)11U };

            applicationNonVisualDrawingProperties37.Append(placeholderShape28);

            nonVisualShapeProperties28.Append(nonVisualDrawingProperties37);
            nonVisualShapeProperties28.Append(nonVisualShapeDrawingProperties28);
            nonVisualShapeProperties28.Append(applicationNonVisualDrawingProperties37);
            ShapeProperties shapeProperties29 = new ShapeProperties();

            TextBody textBody28 = new TextBody();
            A.BodyProperties bodyProperties28 = new A.BodyProperties();
            A.ListStyle listStyle28 = new A.ListStyle();

            A.Paragraph paragraph40 = new A.Paragraph();
            A.EndParagraphRunProperties endParagraphRunProperties26 = new A.EndParagraphRunProperties(){ Language = "cs-CZ" };

            paragraph40.Append(endParagraphRunProperties26);

            textBody28.Append(bodyProperties28);
            textBody28.Append(listStyle28);
            textBody28.Append(paragraph40);

            shape28.Append(nonVisualShapeProperties28);
            shape28.Append(shapeProperties29);
            shape28.Append(textBody28);

            Shape shape29 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties29 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties38 = new NonVisualDrawingProperties(){ Id = (UInt32Value)6U, Name = "Slide Number Placeholder 5" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties29 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks29 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties29.Append(shapeLocks29);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties38 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape29 = new PlaceholderShape(){ Type = PlaceholderValues.SlideNumber, Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)12U };

            applicationNonVisualDrawingProperties38.Append(placeholderShape29);

            nonVisualShapeProperties29.Append(nonVisualDrawingProperties38);
            nonVisualShapeProperties29.Append(nonVisualShapeDrawingProperties29);
            nonVisualShapeProperties29.Append(applicationNonVisualDrawingProperties38);
            ShapeProperties shapeProperties30 = new ShapeProperties();

            TextBody textBody29 = new TextBody();
            A.BodyProperties bodyProperties29 = new A.BodyProperties();
            A.ListStyle listStyle29 = new A.ListStyle();

            A.Paragraph paragraph41 = new A.Paragraph();

            A.Field field12 = new A.Field(){ Id = "{5D7F2B33-3E35-4F97-A448-9FEB151DF743}", Type = "slidenum" };

            A.RunProperties runProperties35 = new A.RunProperties(){ Language = "cs-CZ" };
            runProperties35.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            A.Text text35 = new A.Text();
            text35.Text = "‹#›";

            field12.Append(runProperties35);
            field12.Append(text35);
            A.EndParagraphRunProperties endParagraphRunProperties27 = new A.EndParagraphRunProperties(){ Language = "cs-CZ" };

            paragraph41.Append(field12);
            paragraph41.Append(endParagraphRunProperties27);

            textBody29.Append(bodyProperties29);
            textBody29.Append(listStyle29);
            textBody29.Append(paragraph41);

            shape29.Append(nonVisualShapeProperties29);
            shape29.Append(shapeProperties30);
            shape29.Append(textBody29);

            shapeTree7.Append(nonVisualGroupShapeProperties7);
            shapeTree7.Append(groupShapeProperties7);
            shapeTree7.Append(shape25);
            shapeTree7.Append(shape26);
            shapeTree7.Append(shape27);
            shapeTree7.Append(shape28);
            shapeTree7.Append(shape29);

            CommonSlideDataExtensionList commonSlideDataExtensionList7 = new CommonSlideDataExtensionList();

            CommonSlideDataExtension commonSlideDataExtension7 = new CommonSlideDataExtension(){ Uri = "{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}" };

            P14.CreationId creationId7 = new P14.CreationId(){ Val = (UInt32Value)4126228096U };
            creationId7.AddNamespaceDeclaration("p14", "http://schemas.microsoft.com/office/powerpoint/2010/main");

            commonSlideDataExtension7.Append(creationId7);

            commonSlideDataExtensionList7.Append(commonSlideDataExtension7);

            commonSlideData7.Append(shapeTree7);
            commonSlideData7.Append(commonSlideDataExtensionList7);

            ColorMapOverride colorMapOverride6 = new ColorMapOverride();
            A.MasterColorMapping masterColorMapping6 = new A.MasterColorMapping();

            colorMapOverride6.Append(masterColorMapping6);

            slideLayout5.Append(commonSlideData7);
            slideLayout5.Append(colorMapOverride6);

            slideLayoutPart5.SlideLayout = slideLayout5;
        }

        // Generates content of slideLayoutPart6.
        private void GenerateSlideLayoutPart6Content(SlideLayoutPart slideLayoutPart6)
        {
            SlideLayout slideLayout6 = new SlideLayout(){ Type = SlideLayoutValues.TitleOnly, Preserve = true };
            slideLayout6.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            slideLayout6.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            slideLayout6.AddNamespaceDeclaration("p", "http://schemas.openxmlformats.org/presentationml/2006/main");

            CommonSlideData commonSlideData8 = new CommonSlideData(){ Name = "Jenom nadpis" };

            ShapeTree shapeTree8 = new ShapeTree();

            NonVisualGroupShapeProperties nonVisualGroupShapeProperties8 = new NonVisualGroupShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties39 = new NonVisualDrawingProperties(){ Id = (UInt32Value)1U, Name = "" };
            NonVisualGroupShapeDrawingProperties nonVisualGroupShapeDrawingProperties8 = new NonVisualGroupShapeDrawingProperties();
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties39 = new ApplicationNonVisualDrawingProperties();

            nonVisualGroupShapeProperties8.Append(nonVisualDrawingProperties39);
            nonVisualGroupShapeProperties8.Append(nonVisualGroupShapeDrawingProperties8);
            nonVisualGroupShapeProperties8.Append(applicationNonVisualDrawingProperties39);

            GroupShapeProperties groupShapeProperties8 = new GroupShapeProperties();

            A.TransformGroup transformGroup8 = new A.TransformGroup();
            A.Offset offset22 = new A.Offset(){ X = 0L, Y = 0L };
            A.Extents extents22 = new A.Extents(){ Cx = 0L, Cy = 0L };
            A.ChildOffset childOffset8 = new A.ChildOffset(){ X = 0L, Y = 0L };
            A.ChildExtents childExtents8 = new A.ChildExtents(){ Cx = 0L, Cy = 0L };

            transformGroup8.Append(offset22);
            transformGroup8.Append(extents22);
            transformGroup8.Append(childOffset8);
            transformGroup8.Append(childExtents8);

            groupShapeProperties8.Append(transformGroup8);

            Shape shape30 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties30 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties40 = new NonVisualDrawingProperties(){ Id = (UInt32Value)2U, Name = "Title 1" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties30 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks30 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties30.Append(shapeLocks30);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties40 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape30 = new PlaceholderShape(){ Type = PlaceholderValues.Title };

            applicationNonVisualDrawingProperties40.Append(placeholderShape30);

            nonVisualShapeProperties30.Append(nonVisualDrawingProperties40);
            nonVisualShapeProperties30.Append(nonVisualShapeDrawingProperties30);
            nonVisualShapeProperties30.Append(applicationNonVisualDrawingProperties40);
            ShapeProperties shapeProperties31 = new ShapeProperties();

            TextBody textBody30 = new TextBody();
            A.BodyProperties bodyProperties30 = new A.BodyProperties();
            A.ListStyle listStyle30 = new A.ListStyle();

            A.Paragraph paragraph42 = new A.Paragraph();

            A.Run run24 = new A.Run();
            A.RunProperties runProperties36 = new A.RunProperties(){ Language = "cs-CZ" };
            A.Text text36 = new A.Text();
            text36.Text = "Kliknutím lze upravit styl.";

            run24.Append(runProperties36);
            run24.Append(text36);
            A.EndParagraphRunProperties endParagraphRunProperties28 = new A.EndParagraphRunProperties(){ Language = "en-US", Dirty = false };

            paragraph42.Append(run24);
            paragraph42.Append(endParagraphRunProperties28);

            textBody30.Append(bodyProperties30);
            textBody30.Append(listStyle30);
            textBody30.Append(paragraph42);

            shape30.Append(nonVisualShapeProperties30);
            shape30.Append(shapeProperties31);
            shape30.Append(textBody30);

            Shape shape31 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties31 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties41 = new NonVisualDrawingProperties(){ Id = (UInt32Value)3U, Name = "Date Placeholder 2" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties31 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks31 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties31.Append(shapeLocks31);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties41 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape31 = new PlaceholderShape(){ Type = PlaceholderValues.DateAndTime, Size = PlaceholderSizeValues.Half, Index = (UInt32Value)10U };

            applicationNonVisualDrawingProperties41.Append(placeholderShape31);

            nonVisualShapeProperties31.Append(nonVisualDrawingProperties41);
            nonVisualShapeProperties31.Append(nonVisualShapeDrawingProperties31);
            nonVisualShapeProperties31.Append(applicationNonVisualDrawingProperties41);
            ShapeProperties shapeProperties32 = new ShapeProperties();

            TextBody textBody31 = new TextBody();
            A.BodyProperties bodyProperties31 = new A.BodyProperties();
            A.ListStyle listStyle31 = new A.ListStyle();

            A.Paragraph paragraph43 = new A.Paragraph();

            A.Field field13 = new A.Field(){ Id = "{1D065981-0097-4BA9-A692-9FF652AFAF35}", Type = "datetimeFigureOut" };

            A.RunProperties runProperties37 = new A.RunProperties(){ Language = "cs-CZ" };
            runProperties37.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            A.Text text37 = new A.Text();
            text37.Text = "14.03.2023";

            field13.Append(runProperties37);
            field13.Append(text37);
            A.EndParagraphRunProperties endParagraphRunProperties29 = new A.EndParagraphRunProperties(){ Language = "cs-CZ" };

            paragraph43.Append(field13);
            paragraph43.Append(endParagraphRunProperties29);

            textBody31.Append(bodyProperties31);
            textBody31.Append(listStyle31);
            textBody31.Append(paragraph43);

            shape31.Append(nonVisualShapeProperties31);
            shape31.Append(shapeProperties32);
            shape31.Append(textBody31);

            Shape shape32 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties32 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties42 = new NonVisualDrawingProperties(){ Id = (UInt32Value)4U, Name = "Footer Placeholder 3" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties32 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks32 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties32.Append(shapeLocks32);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties42 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape32 = new PlaceholderShape(){ Type = PlaceholderValues.Footer, Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)11U };

            applicationNonVisualDrawingProperties42.Append(placeholderShape32);

            nonVisualShapeProperties32.Append(nonVisualDrawingProperties42);
            nonVisualShapeProperties32.Append(nonVisualShapeDrawingProperties32);
            nonVisualShapeProperties32.Append(applicationNonVisualDrawingProperties42);
            ShapeProperties shapeProperties33 = new ShapeProperties();

            TextBody textBody32 = new TextBody();
            A.BodyProperties bodyProperties32 = new A.BodyProperties();
            A.ListStyle listStyle32 = new A.ListStyle();

            A.Paragraph paragraph44 = new A.Paragraph();
            A.EndParagraphRunProperties endParagraphRunProperties30 = new A.EndParagraphRunProperties(){ Language = "cs-CZ" };

            paragraph44.Append(endParagraphRunProperties30);

            textBody32.Append(bodyProperties32);
            textBody32.Append(listStyle32);
            textBody32.Append(paragraph44);

            shape32.Append(nonVisualShapeProperties32);
            shape32.Append(shapeProperties33);
            shape32.Append(textBody32);

            Shape shape33 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties33 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties43 = new NonVisualDrawingProperties(){ Id = (UInt32Value)5U, Name = "Slide Number Placeholder 4" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties33 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks33 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties33.Append(shapeLocks33);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties43 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape33 = new PlaceholderShape(){ Type = PlaceholderValues.SlideNumber, Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)12U };

            applicationNonVisualDrawingProperties43.Append(placeholderShape33);

            nonVisualShapeProperties33.Append(nonVisualDrawingProperties43);
            nonVisualShapeProperties33.Append(nonVisualShapeDrawingProperties33);
            nonVisualShapeProperties33.Append(applicationNonVisualDrawingProperties43);
            ShapeProperties shapeProperties34 = new ShapeProperties();

            TextBody textBody33 = new TextBody();
            A.BodyProperties bodyProperties33 = new A.BodyProperties();
            A.ListStyle listStyle33 = new A.ListStyle();

            A.Paragraph paragraph45 = new A.Paragraph();

            A.Field field14 = new A.Field(){ Id = "{5D7F2B33-3E35-4F97-A448-9FEB151DF743}", Type = "slidenum" };

            A.RunProperties runProperties38 = new A.RunProperties(){ Language = "cs-CZ" };
            runProperties38.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            A.Text text38 = new A.Text();
            text38.Text = "‹#›";

            field14.Append(runProperties38);
            field14.Append(text38);
            A.EndParagraphRunProperties endParagraphRunProperties31 = new A.EndParagraphRunProperties(){ Language = "cs-CZ" };

            paragraph45.Append(field14);
            paragraph45.Append(endParagraphRunProperties31);

            textBody33.Append(bodyProperties33);
            textBody33.Append(listStyle33);
            textBody33.Append(paragraph45);

            shape33.Append(nonVisualShapeProperties33);
            shape33.Append(shapeProperties34);
            shape33.Append(textBody33);

            shapeTree8.Append(nonVisualGroupShapeProperties8);
            shapeTree8.Append(groupShapeProperties8);
            shapeTree8.Append(shape30);
            shapeTree8.Append(shape31);
            shapeTree8.Append(shape32);
            shapeTree8.Append(shape33);

            CommonSlideDataExtensionList commonSlideDataExtensionList8 = new CommonSlideDataExtensionList();

            CommonSlideDataExtension commonSlideDataExtension8 = new CommonSlideDataExtension(){ Uri = "{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}" };

            P14.CreationId creationId8 = new P14.CreationId(){ Val = (UInt32Value)1759041178U };
            creationId8.AddNamespaceDeclaration("p14", "http://schemas.microsoft.com/office/powerpoint/2010/main");

            commonSlideDataExtension8.Append(creationId8);

            commonSlideDataExtensionList8.Append(commonSlideDataExtension8);

            commonSlideData8.Append(shapeTree8);
            commonSlideData8.Append(commonSlideDataExtensionList8);

            ColorMapOverride colorMapOverride7 = new ColorMapOverride();
            A.MasterColorMapping masterColorMapping7 = new A.MasterColorMapping();

            colorMapOverride7.Append(masterColorMapping7);

            slideLayout6.Append(commonSlideData8);
            slideLayout6.Append(colorMapOverride7);

            slideLayoutPart6.SlideLayout = slideLayout6;
        }

        // Generates content of slideLayoutPart7.
        private void GenerateSlideLayoutPart7Content(SlideLayoutPart slideLayoutPart7)
        {
            SlideLayout slideLayout7 = new SlideLayout(){ Type = SlideLayoutValues.VerticalTitleAndText, Preserve = true };
            slideLayout7.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            slideLayout7.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            slideLayout7.AddNamespaceDeclaration("p", "http://schemas.openxmlformats.org/presentationml/2006/main");

            CommonSlideData commonSlideData9 = new CommonSlideData(){ Name = "Svislý nadpis a text" };

            ShapeTree shapeTree9 = new ShapeTree();

            NonVisualGroupShapeProperties nonVisualGroupShapeProperties9 = new NonVisualGroupShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties44 = new NonVisualDrawingProperties(){ Id = (UInt32Value)1U, Name = "" };
            NonVisualGroupShapeDrawingProperties nonVisualGroupShapeDrawingProperties9 = new NonVisualGroupShapeDrawingProperties();
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties44 = new ApplicationNonVisualDrawingProperties();

            nonVisualGroupShapeProperties9.Append(nonVisualDrawingProperties44);
            nonVisualGroupShapeProperties9.Append(nonVisualGroupShapeDrawingProperties9);
            nonVisualGroupShapeProperties9.Append(applicationNonVisualDrawingProperties44);

            GroupShapeProperties groupShapeProperties9 = new GroupShapeProperties();

            A.TransformGroup transformGroup9 = new A.TransformGroup();
            A.Offset offset23 = new A.Offset(){ X = 0L, Y = 0L };
            A.Extents extents23 = new A.Extents(){ Cx = 0L, Cy = 0L };
            A.ChildOffset childOffset9 = new A.ChildOffset(){ X = 0L, Y = 0L };
            A.ChildExtents childExtents9 = new A.ChildExtents(){ Cx = 0L, Cy = 0L };

            transformGroup9.Append(offset23);
            transformGroup9.Append(extents23);
            transformGroup9.Append(childOffset9);
            transformGroup9.Append(childExtents9);

            groupShapeProperties9.Append(transformGroup9);

            Shape shape34 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties34 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties45 = new NonVisualDrawingProperties(){ Id = (UInt32Value)2U, Name = "Vertical Title 1" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties34 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks34 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties34.Append(shapeLocks34);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties45 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape34 = new PlaceholderShape(){ Type = PlaceholderValues.Title, Orientation = DirectionValues.Vertical };

            applicationNonVisualDrawingProperties45.Append(placeholderShape34);

            nonVisualShapeProperties34.Append(nonVisualDrawingProperties45);
            nonVisualShapeProperties34.Append(nonVisualShapeDrawingProperties34);
            nonVisualShapeProperties34.Append(applicationNonVisualDrawingProperties45);

            ShapeProperties shapeProperties35 = new ShapeProperties();

            A.Transform2D transform2D14 = new A.Transform2D();
            A.Offset offset24 = new A.Offset(){ X = 6543675L, Y = 365125L };
            A.Extents extents24 = new A.Extents(){ Cx = 1971675L, Cy = 5811838L };

            transform2D14.Append(offset24);
            transform2D14.Append(extents24);

            shapeProperties35.Append(transform2D14);

            TextBody textBody34 = new TextBody();
            A.BodyProperties bodyProperties34 = new A.BodyProperties(){ Vertical = A.TextVerticalValues.EastAsianVetical };
            A.ListStyle listStyle34 = new A.ListStyle();

            A.Paragraph paragraph46 = new A.Paragraph();

            A.Run run25 = new A.Run();
            A.RunProperties runProperties39 = new A.RunProperties(){ Language = "cs-CZ" };
            A.Text text39 = new A.Text();
            text39.Text = "Kliknutím lze upravit styl.";

            run25.Append(runProperties39);
            run25.Append(text39);
            A.EndParagraphRunProperties endParagraphRunProperties32 = new A.EndParagraphRunProperties(){ Language = "en-US", Dirty = false };

            paragraph46.Append(run25);
            paragraph46.Append(endParagraphRunProperties32);

            textBody34.Append(bodyProperties34);
            textBody34.Append(listStyle34);
            textBody34.Append(paragraph46);

            shape34.Append(nonVisualShapeProperties34);
            shape34.Append(shapeProperties35);
            shape34.Append(textBody34);

            Shape shape35 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties35 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties46 = new NonVisualDrawingProperties(){ Id = (UInt32Value)3U, Name = "Vertical Text Placeholder 2" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties35 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks35 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties35.Append(shapeLocks35);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties46 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape35 = new PlaceholderShape(){ Type = PlaceholderValues.Body, Orientation = DirectionValues.Vertical, Index = (UInt32Value)1U };

            applicationNonVisualDrawingProperties46.Append(placeholderShape35);

            nonVisualShapeProperties35.Append(nonVisualDrawingProperties46);
            nonVisualShapeProperties35.Append(nonVisualShapeDrawingProperties35);
            nonVisualShapeProperties35.Append(applicationNonVisualDrawingProperties46);

            ShapeProperties shapeProperties36 = new ShapeProperties();

            A.Transform2D transform2D15 = new A.Transform2D();
            A.Offset offset25 = new A.Offset(){ X = 628650L, Y = 365125L };
            A.Extents extents25 = new A.Extents(){ Cx = 5800725L, Cy = 5811838L };

            transform2D15.Append(offset25);
            transform2D15.Append(extents25);

            shapeProperties36.Append(transform2D15);

            TextBody textBody35 = new TextBody();
            A.BodyProperties bodyProperties35 = new A.BodyProperties(){ Vertical = A.TextVerticalValues.EastAsianVetical };
            A.ListStyle listStyle35 = new A.ListStyle();

            A.Paragraph paragraph47 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties18 = new A.ParagraphProperties(){ Level = 0 };

            A.Run run26 = new A.Run();
            A.RunProperties runProperties40 = new A.RunProperties(){ Language = "cs-CZ" };
            A.Text text40 = new A.Text();
            text40.Text = "Po kliknutí můžete upravovat styly textu v předloze.";

            run26.Append(runProperties40);
            run26.Append(text40);

            paragraph47.Append(paragraphProperties18);
            paragraph47.Append(run26);

            A.Paragraph paragraph48 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties19 = new A.ParagraphProperties(){ Level = 1 };

            A.Run run27 = new A.Run();
            A.RunProperties runProperties41 = new A.RunProperties(){ Language = "cs-CZ" };
            A.Text text41 = new A.Text();
            text41.Text = "Druhá úroveň";

            run27.Append(runProperties41);
            run27.Append(text41);

            paragraph48.Append(paragraphProperties19);
            paragraph48.Append(run27);

            A.Paragraph paragraph49 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties20 = new A.ParagraphProperties(){ Level = 2 };

            A.Run run28 = new A.Run();
            A.RunProperties runProperties42 = new A.RunProperties(){ Language = "cs-CZ" };
            A.Text text42 = new A.Text();
            text42.Text = "Třetí úroveň";

            run28.Append(runProperties42);
            run28.Append(text42);

            paragraph49.Append(paragraphProperties20);
            paragraph49.Append(run28);

            A.Paragraph paragraph50 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties21 = new A.ParagraphProperties(){ Level = 3 };

            A.Run run29 = new A.Run();
            A.RunProperties runProperties43 = new A.RunProperties(){ Language = "cs-CZ" };
            A.Text text43 = new A.Text();
            text43.Text = "Čtvrtá úroveň";

            run29.Append(runProperties43);
            run29.Append(text43);

            paragraph50.Append(paragraphProperties21);
            paragraph50.Append(run29);

            A.Paragraph paragraph51 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties22 = new A.ParagraphProperties(){ Level = 4 };

            A.Run run30 = new A.Run();
            A.RunProperties runProperties44 = new A.RunProperties(){ Language = "cs-CZ" };
            A.Text text44 = new A.Text();
            text44.Text = "Pátá úroveň";

            run30.Append(runProperties44);
            run30.Append(text44);
            A.EndParagraphRunProperties endParagraphRunProperties33 = new A.EndParagraphRunProperties(){ Language = "en-US", Dirty = false };

            paragraph51.Append(paragraphProperties22);
            paragraph51.Append(run30);
            paragraph51.Append(endParagraphRunProperties33);

            textBody35.Append(bodyProperties35);
            textBody35.Append(listStyle35);
            textBody35.Append(paragraph47);
            textBody35.Append(paragraph48);
            textBody35.Append(paragraph49);
            textBody35.Append(paragraph50);
            textBody35.Append(paragraph51);

            shape35.Append(nonVisualShapeProperties35);
            shape35.Append(shapeProperties36);
            shape35.Append(textBody35);

            Shape shape36 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties36 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties47 = new NonVisualDrawingProperties(){ Id = (UInt32Value)4U, Name = "Date Placeholder 3" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties36 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks36 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties36.Append(shapeLocks36);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties47 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape36 = new PlaceholderShape(){ Type = PlaceholderValues.DateAndTime, Size = PlaceholderSizeValues.Half, Index = (UInt32Value)10U };

            applicationNonVisualDrawingProperties47.Append(placeholderShape36);

            nonVisualShapeProperties36.Append(nonVisualDrawingProperties47);
            nonVisualShapeProperties36.Append(nonVisualShapeDrawingProperties36);
            nonVisualShapeProperties36.Append(applicationNonVisualDrawingProperties47);
            ShapeProperties shapeProperties37 = new ShapeProperties();

            TextBody textBody36 = new TextBody();
            A.BodyProperties bodyProperties36 = new A.BodyProperties();
            A.ListStyle listStyle36 = new A.ListStyle();

            A.Paragraph paragraph52 = new A.Paragraph();

            A.Field field15 = new A.Field(){ Id = "{1D065981-0097-4BA9-A692-9FF652AFAF35}", Type = "datetimeFigureOut" };

            A.RunProperties runProperties45 = new A.RunProperties(){ Language = "cs-CZ" };
            runProperties45.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            A.Text text45 = new A.Text();
            text45.Text = "14.03.2023";

            field15.Append(runProperties45);
            field15.Append(text45);
            A.EndParagraphRunProperties endParagraphRunProperties34 = new A.EndParagraphRunProperties(){ Language = "cs-CZ" };

            paragraph52.Append(field15);
            paragraph52.Append(endParagraphRunProperties34);

            textBody36.Append(bodyProperties36);
            textBody36.Append(listStyle36);
            textBody36.Append(paragraph52);

            shape36.Append(nonVisualShapeProperties36);
            shape36.Append(shapeProperties37);
            shape36.Append(textBody36);

            Shape shape37 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties37 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties48 = new NonVisualDrawingProperties(){ Id = (UInt32Value)5U, Name = "Footer Placeholder 4" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties37 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks37 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties37.Append(shapeLocks37);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties48 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape37 = new PlaceholderShape(){ Type = PlaceholderValues.Footer, Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)11U };

            applicationNonVisualDrawingProperties48.Append(placeholderShape37);

            nonVisualShapeProperties37.Append(nonVisualDrawingProperties48);
            nonVisualShapeProperties37.Append(nonVisualShapeDrawingProperties37);
            nonVisualShapeProperties37.Append(applicationNonVisualDrawingProperties48);
            ShapeProperties shapeProperties38 = new ShapeProperties();

            TextBody textBody37 = new TextBody();
            A.BodyProperties bodyProperties37 = new A.BodyProperties();
            A.ListStyle listStyle37 = new A.ListStyle();

            A.Paragraph paragraph53 = new A.Paragraph();
            A.EndParagraphRunProperties endParagraphRunProperties35 = new A.EndParagraphRunProperties(){ Language = "cs-CZ" };

            paragraph53.Append(endParagraphRunProperties35);

            textBody37.Append(bodyProperties37);
            textBody37.Append(listStyle37);
            textBody37.Append(paragraph53);

            shape37.Append(nonVisualShapeProperties37);
            shape37.Append(shapeProperties38);
            shape37.Append(textBody37);

            Shape shape38 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties38 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties49 = new NonVisualDrawingProperties(){ Id = (UInt32Value)6U, Name = "Slide Number Placeholder 5" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties38 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks38 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties38.Append(shapeLocks38);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties49 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape38 = new PlaceholderShape(){ Type = PlaceholderValues.SlideNumber, Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)12U };

            applicationNonVisualDrawingProperties49.Append(placeholderShape38);

            nonVisualShapeProperties38.Append(nonVisualDrawingProperties49);
            nonVisualShapeProperties38.Append(nonVisualShapeDrawingProperties38);
            nonVisualShapeProperties38.Append(applicationNonVisualDrawingProperties49);
            ShapeProperties shapeProperties39 = new ShapeProperties();

            TextBody textBody38 = new TextBody();
            A.BodyProperties bodyProperties38 = new A.BodyProperties();
            A.ListStyle listStyle38 = new A.ListStyle();

            A.Paragraph paragraph54 = new A.Paragraph();

            A.Field field16 = new A.Field(){ Id = "{5D7F2B33-3E35-4F97-A448-9FEB151DF743}", Type = "slidenum" };

            A.RunProperties runProperties46 = new A.RunProperties(){ Language = "cs-CZ" };
            runProperties46.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            A.Text text46 = new A.Text();
            text46.Text = "‹#›";

            field16.Append(runProperties46);
            field16.Append(text46);
            A.EndParagraphRunProperties endParagraphRunProperties36 = new A.EndParagraphRunProperties(){ Language = "cs-CZ" };

            paragraph54.Append(field16);
            paragraph54.Append(endParagraphRunProperties36);

            textBody38.Append(bodyProperties38);
            textBody38.Append(listStyle38);
            textBody38.Append(paragraph54);

            shape38.Append(nonVisualShapeProperties38);
            shape38.Append(shapeProperties39);
            shape38.Append(textBody38);

            shapeTree9.Append(nonVisualGroupShapeProperties9);
            shapeTree9.Append(groupShapeProperties9);
            shapeTree9.Append(shape34);
            shapeTree9.Append(shape35);
            shapeTree9.Append(shape36);
            shapeTree9.Append(shape37);
            shapeTree9.Append(shape38);

            CommonSlideDataExtensionList commonSlideDataExtensionList9 = new CommonSlideDataExtensionList();

            CommonSlideDataExtension commonSlideDataExtension9 = new CommonSlideDataExtension(){ Uri = "{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}" };

            P14.CreationId creationId9 = new P14.CreationId(){ Val = (UInt32Value)3086997787U };
            creationId9.AddNamespaceDeclaration("p14", "http://schemas.microsoft.com/office/powerpoint/2010/main");

            commonSlideDataExtension9.Append(creationId9);

            commonSlideDataExtensionList9.Append(commonSlideDataExtension9);

            commonSlideData9.Append(shapeTree9);
            commonSlideData9.Append(commonSlideDataExtensionList9);

            ColorMapOverride colorMapOverride8 = new ColorMapOverride();
            A.MasterColorMapping masterColorMapping8 = new A.MasterColorMapping();

            colorMapOverride8.Append(masterColorMapping8);

            slideLayout7.Append(commonSlideData9);
            slideLayout7.Append(colorMapOverride8);

            slideLayoutPart7.SlideLayout = slideLayout7;
        }

        // Generates content of slideLayoutPart8.
        private void GenerateSlideLayoutPart8Content(SlideLayoutPart slideLayoutPart8)
        {
            SlideLayout slideLayout8 = new SlideLayout(){ Type = SlideLayoutValues.TwoTextAndTwoObjects, Preserve = true };
            slideLayout8.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            slideLayout8.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            slideLayout8.AddNamespaceDeclaration("p", "http://schemas.openxmlformats.org/presentationml/2006/main");

            CommonSlideData commonSlideData10 = new CommonSlideData(){ Name = "Porovnání" };

            ShapeTree shapeTree10 = new ShapeTree();

            NonVisualGroupShapeProperties nonVisualGroupShapeProperties10 = new NonVisualGroupShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties50 = new NonVisualDrawingProperties(){ Id = (UInt32Value)1U, Name = "" };
            NonVisualGroupShapeDrawingProperties nonVisualGroupShapeDrawingProperties10 = new NonVisualGroupShapeDrawingProperties();
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties50 = new ApplicationNonVisualDrawingProperties();

            nonVisualGroupShapeProperties10.Append(nonVisualDrawingProperties50);
            nonVisualGroupShapeProperties10.Append(nonVisualGroupShapeDrawingProperties10);
            nonVisualGroupShapeProperties10.Append(applicationNonVisualDrawingProperties50);

            GroupShapeProperties groupShapeProperties10 = new GroupShapeProperties();

            A.TransformGroup transformGroup10 = new A.TransformGroup();
            A.Offset offset26 = new A.Offset(){ X = 0L, Y = 0L };
            A.Extents extents26 = new A.Extents(){ Cx = 0L, Cy = 0L };
            A.ChildOffset childOffset10 = new A.ChildOffset(){ X = 0L, Y = 0L };
            A.ChildExtents childExtents10 = new A.ChildExtents(){ Cx = 0L, Cy = 0L };

            transformGroup10.Append(offset26);
            transformGroup10.Append(extents26);
            transformGroup10.Append(childOffset10);
            transformGroup10.Append(childExtents10);

            groupShapeProperties10.Append(transformGroup10);

            Shape shape39 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties39 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties51 = new NonVisualDrawingProperties(){ Id = (UInt32Value)2U, Name = "Title 1" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties39 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks39 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties39.Append(shapeLocks39);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties51 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape39 = new PlaceholderShape(){ Type = PlaceholderValues.Title };

            applicationNonVisualDrawingProperties51.Append(placeholderShape39);

            nonVisualShapeProperties39.Append(nonVisualDrawingProperties51);
            nonVisualShapeProperties39.Append(nonVisualShapeDrawingProperties39);
            nonVisualShapeProperties39.Append(applicationNonVisualDrawingProperties51);

            ShapeProperties shapeProperties40 = new ShapeProperties();

            A.Transform2D transform2D16 = new A.Transform2D();
            A.Offset offset27 = new A.Offset(){ X = 629841L, Y = 365126L };
            A.Extents extents27 = new A.Extents(){ Cx = 7886700L, Cy = 1325563L };

            transform2D16.Append(offset27);
            transform2D16.Append(extents27);

            shapeProperties40.Append(transform2D16);

            TextBody textBody39 = new TextBody();
            A.BodyProperties bodyProperties39 = new A.BodyProperties();
            A.ListStyle listStyle39 = new A.ListStyle();

            A.Paragraph paragraph55 = new A.Paragraph();

            A.Run run31 = new A.Run();
            A.RunProperties runProperties47 = new A.RunProperties(){ Language = "cs-CZ" };
            A.Text text47 = new A.Text();
            text47.Text = "Kliknutím lze upravit styl.";

            run31.Append(runProperties47);
            run31.Append(text47);
            A.EndParagraphRunProperties endParagraphRunProperties37 = new A.EndParagraphRunProperties(){ Language = "en-US", Dirty = false };

            paragraph55.Append(run31);
            paragraph55.Append(endParagraphRunProperties37);

            textBody39.Append(bodyProperties39);
            textBody39.Append(listStyle39);
            textBody39.Append(paragraph55);

            shape39.Append(nonVisualShapeProperties39);
            shape39.Append(shapeProperties40);
            shape39.Append(textBody39);

            Shape shape40 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties40 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties52 = new NonVisualDrawingProperties(){ Id = (UInt32Value)3U, Name = "Text Placeholder 2" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties40 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks40 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties40.Append(shapeLocks40);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties52 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape40 = new PlaceholderShape(){ Type = PlaceholderValues.Body, Index = (UInt32Value)1U };

            applicationNonVisualDrawingProperties52.Append(placeholderShape40);

            nonVisualShapeProperties40.Append(nonVisualDrawingProperties52);
            nonVisualShapeProperties40.Append(nonVisualShapeDrawingProperties40);
            nonVisualShapeProperties40.Append(applicationNonVisualDrawingProperties52);

            ShapeProperties shapeProperties41 = new ShapeProperties();

            A.Transform2D transform2D17 = new A.Transform2D();
            A.Offset offset28 = new A.Offset(){ X = 629842L, Y = 1681163L };
            A.Extents extents28 = new A.Extents(){ Cx = 3868340L, Cy = 823912L };

            transform2D17.Append(offset28);
            transform2D17.Append(extents28);

            shapeProperties41.Append(transform2D17);

            TextBody textBody40 = new TextBody();
            A.BodyProperties bodyProperties40 = new A.BodyProperties(){ Anchor = A.TextAnchoringTypeValues.Bottom };

            A.ListStyle listStyle40 = new A.ListStyle();

            A.Level1ParagraphProperties level1ParagraphProperties15 = new A.Level1ParagraphProperties(){ LeftMargin = 0, Indent = 0 };
            A.NoBullet noBullet29 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties73 = new A.DefaultRunProperties(){ FontSize = 2400, Bold = true };

            level1ParagraphProperties15.Append(noBullet29);
            level1ParagraphProperties15.Append(defaultRunProperties73);

            A.Level2ParagraphProperties level2ParagraphProperties8 = new A.Level2ParagraphProperties(){ LeftMargin = 457200, Indent = 0 };
            A.NoBullet noBullet30 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties74 = new A.DefaultRunProperties(){ FontSize = 2000, Bold = true };

            level2ParagraphProperties8.Append(noBullet30);
            level2ParagraphProperties8.Append(defaultRunProperties74);

            A.Level3ParagraphProperties level3ParagraphProperties8 = new A.Level3ParagraphProperties(){ LeftMargin = 914400, Indent = 0 };
            A.NoBullet noBullet31 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties75 = new A.DefaultRunProperties(){ FontSize = 1800, Bold = true };

            level3ParagraphProperties8.Append(noBullet31);
            level3ParagraphProperties8.Append(defaultRunProperties75);

            A.Level4ParagraphProperties level4ParagraphProperties8 = new A.Level4ParagraphProperties(){ LeftMargin = 1371600, Indent = 0 };
            A.NoBullet noBullet32 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties76 = new A.DefaultRunProperties(){ FontSize = 1600, Bold = true };

            level4ParagraphProperties8.Append(noBullet32);
            level4ParagraphProperties8.Append(defaultRunProperties76);

            A.Level5ParagraphProperties level5ParagraphProperties8 = new A.Level5ParagraphProperties(){ LeftMargin = 1828800, Indent = 0 };
            A.NoBullet noBullet33 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties77 = new A.DefaultRunProperties(){ FontSize = 1600, Bold = true };

            level5ParagraphProperties8.Append(noBullet33);
            level5ParagraphProperties8.Append(defaultRunProperties77);

            A.Level6ParagraphProperties level6ParagraphProperties8 = new A.Level6ParagraphProperties(){ LeftMargin = 2286000, Indent = 0 };
            A.NoBullet noBullet34 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties78 = new A.DefaultRunProperties(){ FontSize = 1600, Bold = true };

            level6ParagraphProperties8.Append(noBullet34);
            level6ParagraphProperties8.Append(defaultRunProperties78);

            A.Level7ParagraphProperties level7ParagraphProperties8 = new A.Level7ParagraphProperties(){ LeftMargin = 2743200, Indent = 0 };
            A.NoBullet noBullet35 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties79 = new A.DefaultRunProperties(){ FontSize = 1600, Bold = true };

            level7ParagraphProperties8.Append(noBullet35);
            level7ParagraphProperties8.Append(defaultRunProperties79);

            A.Level8ParagraphProperties level8ParagraphProperties8 = new A.Level8ParagraphProperties(){ LeftMargin = 3200400, Indent = 0 };
            A.NoBullet noBullet36 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties80 = new A.DefaultRunProperties(){ FontSize = 1600, Bold = true };

            level8ParagraphProperties8.Append(noBullet36);
            level8ParagraphProperties8.Append(defaultRunProperties80);

            A.Level9ParagraphProperties level9ParagraphProperties8 = new A.Level9ParagraphProperties(){ LeftMargin = 3657600, Indent = 0 };
            A.NoBullet noBullet37 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties81 = new A.DefaultRunProperties(){ FontSize = 1600, Bold = true };

            level9ParagraphProperties8.Append(noBullet37);
            level9ParagraphProperties8.Append(defaultRunProperties81);

            listStyle40.Append(level1ParagraphProperties15);
            listStyle40.Append(level2ParagraphProperties8);
            listStyle40.Append(level3ParagraphProperties8);
            listStyle40.Append(level4ParagraphProperties8);
            listStyle40.Append(level5ParagraphProperties8);
            listStyle40.Append(level6ParagraphProperties8);
            listStyle40.Append(level7ParagraphProperties8);
            listStyle40.Append(level8ParagraphProperties8);
            listStyle40.Append(level9ParagraphProperties8);

            A.Paragraph paragraph56 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties23 = new A.ParagraphProperties(){ Level = 0 };

            A.Run run32 = new A.Run();
            A.RunProperties runProperties48 = new A.RunProperties(){ Language = "cs-CZ" };
            A.Text text48 = new A.Text();
            text48.Text = "Po kliknutí můžete upravovat styly textu v předloze.";

            run32.Append(runProperties48);
            run32.Append(text48);

            paragraph56.Append(paragraphProperties23);
            paragraph56.Append(run32);

            textBody40.Append(bodyProperties40);
            textBody40.Append(listStyle40);
            textBody40.Append(paragraph56);

            shape40.Append(nonVisualShapeProperties40);
            shape40.Append(shapeProperties41);
            shape40.Append(textBody40);

            Shape shape41 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties41 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties53 = new NonVisualDrawingProperties(){ Id = (UInt32Value)4U, Name = "Content Placeholder 3" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties41 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks41 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties41.Append(shapeLocks41);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties53 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape41 = new PlaceholderShape(){ Size = PlaceholderSizeValues.Half, Index = (UInt32Value)2U };

            applicationNonVisualDrawingProperties53.Append(placeholderShape41);

            nonVisualShapeProperties41.Append(nonVisualDrawingProperties53);
            nonVisualShapeProperties41.Append(nonVisualShapeDrawingProperties41);
            nonVisualShapeProperties41.Append(applicationNonVisualDrawingProperties53);

            ShapeProperties shapeProperties42 = new ShapeProperties();

            A.Transform2D transform2D18 = new A.Transform2D();
            A.Offset offset29 = new A.Offset(){ X = 629842L, Y = 2505075L };
            A.Extents extents29 = new A.Extents(){ Cx = 3868340L, Cy = 3684588L };

            transform2D18.Append(offset29);
            transform2D18.Append(extents29);

            shapeProperties42.Append(transform2D18);

            TextBody textBody41 = new TextBody();
            A.BodyProperties bodyProperties41 = new A.BodyProperties();
            A.ListStyle listStyle41 = new A.ListStyle();

            A.Paragraph paragraph57 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties24 = new A.ParagraphProperties(){ Level = 0 };

            A.Run run33 = new A.Run();
            A.RunProperties runProperties49 = new A.RunProperties(){ Language = "cs-CZ" };
            A.Text text49 = new A.Text();
            text49.Text = "Po kliknutí můžete upravovat styly textu v předloze.";

            run33.Append(runProperties49);
            run33.Append(text49);

            paragraph57.Append(paragraphProperties24);
            paragraph57.Append(run33);

            A.Paragraph paragraph58 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties25 = new A.ParagraphProperties(){ Level = 1 };

            A.Run run34 = new A.Run();
            A.RunProperties runProperties50 = new A.RunProperties(){ Language = "cs-CZ" };
            A.Text text50 = new A.Text();
            text50.Text = "Druhá úroveň";

            run34.Append(runProperties50);
            run34.Append(text50);

            paragraph58.Append(paragraphProperties25);
            paragraph58.Append(run34);

            A.Paragraph paragraph59 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties26 = new A.ParagraphProperties(){ Level = 2 };

            A.Run run35 = new A.Run();
            A.RunProperties runProperties51 = new A.RunProperties(){ Language = "cs-CZ" };
            A.Text text51 = new A.Text();
            text51.Text = "Třetí úroveň";

            run35.Append(runProperties51);
            run35.Append(text51);

            paragraph59.Append(paragraphProperties26);
            paragraph59.Append(run35);

            A.Paragraph paragraph60 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties27 = new A.ParagraphProperties(){ Level = 3 };

            A.Run run36 = new A.Run();
            A.RunProperties runProperties52 = new A.RunProperties(){ Language = "cs-CZ" };
            A.Text text52 = new A.Text();
            text52.Text = "Čtvrtá úroveň";

            run36.Append(runProperties52);
            run36.Append(text52);

            paragraph60.Append(paragraphProperties27);
            paragraph60.Append(run36);

            A.Paragraph paragraph61 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties28 = new A.ParagraphProperties(){ Level = 4 };

            A.Run run37 = new A.Run();
            A.RunProperties runProperties53 = new A.RunProperties(){ Language = "cs-CZ" };
            A.Text text53 = new A.Text();
            text53.Text = "Pátá úroveň";

            run37.Append(runProperties53);
            run37.Append(text53);
            A.EndParagraphRunProperties endParagraphRunProperties38 = new A.EndParagraphRunProperties(){ Language = "en-US", Dirty = false };

            paragraph61.Append(paragraphProperties28);
            paragraph61.Append(run37);
            paragraph61.Append(endParagraphRunProperties38);

            textBody41.Append(bodyProperties41);
            textBody41.Append(listStyle41);
            textBody41.Append(paragraph57);
            textBody41.Append(paragraph58);
            textBody41.Append(paragraph59);
            textBody41.Append(paragraph60);
            textBody41.Append(paragraph61);

            shape41.Append(nonVisualShapeProperties41);
            shape41.Append(shapeProperties42);
            shape41.Append(textBody41);

            Shape shape42 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties42 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties54 = new NonVisualDrawingProperties(){ Id = (UInt32Value)5U, Name = "Text Placeholder 4" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties42 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks42 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties42.Append(shapeLocks42);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties54 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape42 = new PlaceholderShape(){ Type = PlaceholderValues.Body, Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)3U };

            applicationNonVisualDrawingProperties54.Append(placeholderShape42);

            nonVisualShapeProperties42.Append(nonVisualDrawingProperties54);
            nonVisualShapeProperties42.Append(nonVisualShapeDrawingProperties42);
            nonVisualShapeProperties42.Append(applicationNonVisualDrawingProperties54);

            ShapeProperties shapeProperties43 = new ShapeProperties();

            A.Transform2D transform2D19 = new A.Transform2D();
            A.Offset offset30 = new A.Offset(){ X = 4629150L, Y = 1681163L };
            A.Extents extents30 = new A.Extents(){ Cx = 3887391L, Cy = 823912L };

            transform2D19.Append(offset30);
            transform2D19.Append(extents30);

            shapeProperties43.Append(transform2D19);

            TextBody textBody42 = new TextBody();
            A.BodyProperties bodyProperties42 = new A.BodyProperties(){ Anchor = A.TextAnchoringTypeValues.Bottom };

            A.ListStyle listStyle42 = new A.ListStyle();

            A.Level1ParagraphProperties level1ParagraphProperties16 = new A.Level1ParagraphProperties(){ LeftMargin = 0, Indent = 0 };
            A.NoBullet noBullet38 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties82 = new A.DefaultRunProperties(){ FontSize = 2400, Bold = true };

            level1ParagraphProperties16.Append(noBullet38);
            level1ParagraphProperties16.Append(defaultRunProperties82);

            A.Level2ParagraphProperties level2ParagraphProperties9 = new A.Level2ParagraphProperties(){ LeftMargin = 457200, Indent = 0 };
            A.NoBullet noBullet39 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties83 = new A.DefaultRunProperties(){ FontSize = 2000, Bold = true };

            level2ParagraphProperties9.Append(noBullet39);
            level2ParagraphProperties9.Append(defaultRunProperties83);

            A.Level3ParagraphProperties level3ParagraphProperties9 = new A.Level3ParagraphProperties(){ LeftMargin = 914400, Indent = 0 };
            A.NoBullet noBullet40 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties84 = new A.DefaultRunProperties(){ FontSize = 1800, Bold = true };

            level3ParagraphProperties9.Append(noBullet40);
            level3ParagraphProperties9.Append(defaultRunProperties84);

            A.Level4ParagraphProperties level4ParagraphProperties9 = new A.Level4ParagraphProperties(){ LeftMargin = 1371600, Indent = 0 };
            A.NoBullet noBullet41 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties85 = new A.DefaultRunProperties(){ FontSize = 1600, Bold = true };

            level4ParagraphProperties9.Append(noBullet41);
            level4ParagraphProperties9.Append(defaultRunProperties85);

            A.Level5ParagraphProperties level5ParagraphProperties9 = new A.Level5ParagraphProperties(){ LeftMargin = 1828800, Indent = 0 };
            A.NoBullet noBullet42 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties86 = new A.DefaultRunProperties(){ FontSize = 1600, Bold = true };

            level5ParagraphProperties9.Append(noBullet42);
            level5ParagraphProperties9.Append(defaultRunProperties86);

            A.Level6ParagraphProperties level6ParagraphProperties9 = new A.Level6ParagraphProperties(){ LeftMargin = 2286000, Indent = 0 };
            A.NoBullet noBullet43 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties87 = new A.DefaultRunProperties(){ FontSize = 1600, Bold = true };

            level6ParagraphProperties9.Append(noBullet43);
            level6ParagraphProperties9.Append(defaultRunProperties87);

            A.Level7ParagraphProperties level7ParagraphProperties9 = new A.Level7ParagraphProperties(){ LeftMargin = 2743200, Indent = 0 };
            A.NoBullet noBullet44 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties88 = new A.DefaultRunProperties(){ FontSize = 1600, Bold = true };

            level7ParagraphProperties9.Append(noBullet44);
            level7ParagraphProperties9.Append(defaultRunProperties88);

            A.Level8ParagraphProperties level8ParagraphProperties9 = new A.Level8ParagraphProperties(){ LeftMargin = 3200400, Indent = 0 };
            A.NoBullet noBullet45 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties89 = new A.DefaultRunProperties(){ FontSize = 1600, Bold = true };

            level8ParagraphProperties9.Append(noBullet45);
            level8ParagraphProperties9.Append(defaultRunProperties89);

            A.Level9ParagraphProperties level9ParagraphProperties9 = new A.Level9ParagraphProperties(){ LeftMargin = 3657600, Indent = 0 };
            A.NoBullet noBullet46 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties90 = new A.DefaultRunProperties(){ FontSize = 1600, Bold = true };

            level9ParagraphProperties9.Append(noBullet46);
            level9ParagraphProperties9.Append(defaultRunProperties90);

            listStyle42.Append(level1ParagraphProperties16);
            listStyle42.Append(level2ParagraphProperties9);
            listStyle42.Append(level3ParagraphProperties9);
            listStyle42.Append(level4ParagraphProperties9);
            listStyle42.Append(level5ParagraphProperties9);
            listStyle42.Append(level6ParagraphProperties9);
            listStyle42.Append(level7ParagraphProperties9);
            listStyle42.Append(level8ParagraphProperties9);
            listStyle42.Append(level9ParagraphProperties9);

            A.Paragraph paragraph62 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties29 = new A.ParagraphProperties(){ Level = 0 };

            A.Run run38 = new A.Run();
            A.RunProperties runProperties54 = new A.RunProperties(){ Language = "cs-CZ" };
            A.Text text54 = new A.Text();
            text54.Text = "Po kliknutí můžete upravovat styly textu v předloze.";

            run38.Append(runProperties54);
            run38.Append(text54);

            paragraph62.Append(paragraphProperties29);
            paragraph62.Append(run38);

            textBody42.Append(bodyProperties42);
            textBody42.Append(listStyle42);
            textBody42.Append(paragraph62);

            shape42.Append(nonVisualShapeProperties42);
            shape42.Append(shapeProperties43);
            shape42.Append(textBody42);

            Shape shape43 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties43 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties55 = new NonVisualDrawingProperties(){ Id = (UInt32Value)6U, Name = "Content Placeholder 5" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties43 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks43 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties43.Append(shapeLocks43);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties55 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape43 = new PlaceholderShape(){ Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)4U };

            applicationNonVisualDrawingProperties55.Append(placeholderShape43);

            nonVisualShapeProperties43.Append(nonVisualDrawingProperties55);
            nonVisualShapeProperties43.Append(nonVisualShapeDrawingProperties43);
            nonVisualShapeProperties43.Append(applicationNonVisualDrawingProperties55);

            ShapeProperties shapeProperties44 = new ShapeProperties();

            A.Transform2D transform2D20 = new A.Transform2D();
            A.Offset offset31 = new A.Offset(){ X = 4629150L, Y = 2505075L };
            A.Extents extents31 = new A.Extents(){ Cx = 3887391L, Cy = 3684588L };

            transform2D20.Append(offset31);
            transform2D20.Append(extents31);

            shapeProperties44.Append(transform2D20);

            TextBody textBody43 = new TextBody();
            A.BodyProperties bodyProperties43 = new A.BodyProperties();
            A.ListStyle listStyle43 = new A.ListStyle();

            A.Paragraph paragraph63 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties30 = new A.ParagraphProperties(){ Level = 0 };

            A.Run run39 = new A.Run();
            A.RunProperties runProperties55 = new A.RunProperties(){ Language = "cs-CZ" };
            A.Text text55 = new A.Text();
            text55.Text = "Po kliknutí můžete upravovat styly textu v předloze.";

            run39.Append(runProperties55);
            run39.Append(text55);

            paragraph63.Append(paragraphProperties30);
            paragraph63.Append(run39);

            A.Paragraph paragraph64 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties31 = new A.ParagraphProperties(){ Level = 1 };

            A.Run run40 = new A.Run();
            A.RunProperties runProperties56 = new A.RunProperties(){ Language = "cs-CZ" };
            A.Text text56 = new A.Text();
            text56.Text = "Druhá úroveň";

            run40.Append(runProperties56);
            run40.Append(text56);

            paragraph64.Append(paragraphProperties31);
            paragraph64.Append(run40);

            A.Paragraph paragraph65 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties32 = new A.ParagraphProperties(){ Level = 2 };

            A.Run run41 = new A.Run();
            A.RunProperties runProperties57 = new A.RunProperties(){ Language = "cs-CZ" };
            A.Text text57 = new A.Text();
            text57.Text = "Třetí úroveň";

            run41.Append(runProperties57);
            run41.Append(text57);

            paragraph65.Append(paragraphProperties32);
            paragraph65.Append(run41);

            A.Paragraph paragraph66 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties33 = new A.ParagraphProperties(){ Level = 3 };

            A.Run run42 = new A.Run();
            A.RunProperties runProperties58 = new A.RunProperties(){ Language = "cs-CZ" };
            A.Text text58 = new A.Text();
            text58.Text = "Čtvrtá úroveň";

            run42.Append(runProperties58);
            run42.Append(text58);

            paragraph66.Append(paragraphProperties33);
            paragraph66.Append(run42);

            A.Paragraph paragraph67 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties34 = new A.ParagraphProperties(){ Level = 4 };

            A.Run run43 = new A.Run();
            A.RunProperties runProperties59 = new A.RunProperties(){ Language = "cs-CZ" };
            A.Text text59 = new A.Text();
            text59.Text = "Pátá úroveň";

            run43.Append(runProperties59);
            run43.Append(text59);
            A.EndParagraphRunProperties endParagraphRunProperties39 = new A.EndParagraphRunProperties(){ Language = "en-US", Dirty = false };

            paragraph67.Append(paragraphProperties34);
            paragraph67.Append(run43);
            paragraph67.Append(endParagraphRunProperties39);

            textBody43.Append(bodyProperties43);
            textBody43.Append(listStyle43);
            textBody43.Append(paragraph63);
            textBody43.Append(paragraph64);
            textBody43.Append(paragraph65);
            textBody43.Append(paragraph66);
            textBody43.Append(paragraph67);

            shape43.Append(nonVisualShapeProperties43);
            shape43.Append(shapeProperties44);
            shape43.Append(textBody43);

            Shape shape44 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties44 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties56 = new NonVisualDrawingProperties(){ Id = (UInt32Value)7U, Name = "Date Placeholder 6" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties44 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks44 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties44.Append(shapeLocks44);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties56 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape44 = new PlaceholderShape(){ Type = PlaceholderValues.DateAndTime, Size = PlaceholderSizeValues.Half, Index = (UInt32Value)10U };

            applicationNonVisualDrawingProperties56.Append(placeholderShape44);

            nonVisualShapeProperties44.Append(nonVisualDrawingProperties56);
            nonVisualShapeProperties44.Append(nonVisualShapeDrawingProperties44);
            nonVisualShapeProperties44.Append(applicationNonVisualDrawingProperties56);
            ShapeProperties shapeProperties45 = new ShapeProperties();

            TextBody textBody44 = new TextBody();
            A.BodyProperties bodyProperties44 = new A.BodyProperties();
            A.ListStyle listStyle44 = new A.ListStyle();

            A.Paragraph paragraph68 = new A.Paragraph();

            A.Field field17 = new A.Field(){ Id = "{1D065981-0097-4BA9-A692-9FF652AFAF35}", Type = "datetimeFigureOut" };

            A.RunProperties runProperties60 = new A.RunProperties(){ Language = "cs-CZ" };
            runProperties60.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            A.Text text60 = new A.Text();
            text60.Text = "14.03.2023";

            field17.Append(runProperties60);
            field17.Append(text60);
            A.EndParagraphRunProperties endParagraphRunProperties40 = new A.EndParagraphRunProperties(){ Language = "cs-CZ" };

            paragraph68.Append(field17);
            paragraph68.Append(endParagraphRunProperties40);

            textBody44.Append(bodyProperties44);
            textBody44.Append(listStyle44);
            textBody44.Append(paragraph68);

            shape44.Append(nonVisualShapeProperties44);
            shape44.Append(shapeProperties45);
            shape44.Append(textBody44);

            Shape shape45 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties45 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties57 = new NonVisualDrawingProperties(){ Id = (UInt32Value)8U, Name = "Footer Placeholder 7" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties45 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks45 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties45.Append(shapeLocks45);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties57 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape45 = new PlaceholderShape(){ Type = PlaceholderValues.Footer, Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)11U };

            applicationNonVisualDrawingProperties57.Append(placeholderShape45);

            nonVisualShapeProperties45.Append(nonVisualDrawingProperties57);
            nonVisualShapeProperties45.Append(nonVisualShapeDrawingProperties45);
            nonVisualShapeProperties45.Append(applicationNonVisualDrawingProperties57);
            ShapeProperties shapeProperties46 = new ShapeProperties();

            TextBody textBody45 = new TextBody();
            A.BodyProperties bodyProperties45 = new A.BodyProperties();
            A.ListStyle listStyle45 = new A.ListStyle();

            A.Paragraph paragraph69 = new A.Paragraph();
            A.EndParagraphRunProperties endParagraphRunProperties41 = new A.EndParagraphRunProperties(){ Language = "cs-CZ" };

            paragraph69.Append(endParagraphRunProperties41);

            textBody45.Append(bodyProperties45);
            textBody45.Append(listStyle45);
            textBody45.Append(paragraph69);

            shape45.Append(nonVisualShapeProperties45);
            shape45.Append(shapeProperties46);
            shape45.Append(textBody45);

            Shape shape46 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties46 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties58 = new NonVisualDrawingProperties(){ Id = (UInt32Value)9U, Name = "Slide Number Placeholder 8" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties46 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks46 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties46.Append(shapeLocks46);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties58 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape46 = new PlaceholderShape(){ Type = PlaceholderValues.SlideNumber, Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)12U };

            applicationNonVisualDrawingProperties58.Append(placeholderShape46);

            nonVisualShapeProperties46.Append(nonVisualDrawingProperties58);
            nonVisualShapeProperties46.Append(nonVisualShapeDrawingProperties46);
            nonVisualShapeProperties46.Append(applicationNonVisualDrawingProperties58);
            ShapeProperties shapeProperties47 = new ShapeProperties();

            TextBody textBody46 = new TextBody();
            A.BodyProperties bodyProperties46 = new A.BodyProperties();
            A.ListStyle listStyle46 = new A.ListStyle();

            A.Paragraph paragraph70 = new A.Paragraph();

            A.Field field18 = new A.Field(){ Id = "{5D7F2B33-3E35-4F97-A448-9FEB151DF743}", Type = "slidenum" };

            A.RunProperties runProperties61 = new A.RunProperties(){ Language = "cs-CZ" };
            runProperties61.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            A.Text text61 = new A.Text();
            text61.Text = "‹#›";

            field18.Append(runProperties61);
            field18.Append(text61);
            A.EndParagraphRunProperties endParagraphRunProperties42 = new A.EndParagraphRunProperties(){ Language = "cs-CZ" };

            paragraph70.Append(field18);
            paragraph70.Append(endParagraphRunProperties42);

            textBody46.Append(bodyProperties46);
            textBody46.Append(listStyle46);
            textBody46.Append(paragraph70);

            shape46.Append(nonVisualShapeProperties46);
            shape46.Append(shapeProperties47);
            shape46.Append(textBody46);

            shapeTree10.Append(nonVisualGroupShapeProperties10);
            shapeTree10.Append(groupShapeProperties10);
            shapeTree10.Append(shape39);
            shapeTree10.Append(shape40);
            shapeTree10.Append(shape41);
            shapeTree10.Append(shape42);
            shapeTree10.Append(shape43);
            shapeTree10.Append(shape44);
            shapeTree10.Append(shape45);
            shapeTree10.Append(shape46);

            CommonSlideDataExtensionList commonSlideDataExtensionList10 = new CommonSlideDataExtensionList();

            CommonSlideDataExtension commonSlideDataExtension10 = new CommonSlideDataExtension(){ Uri = "{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}" };

            P14.CreationId creationId10 = new P14.CreationId(){ Val = (UInt32Value)2069419401U };
            creationId10.AddNamespaceDeclaration("p14", "http://schemas.microsoft.com/office/powerpoint/2010/main");

            commonSlideDataExtension10.Append(creationId10);

            commonSlideDataExtensionList10.Append(commonSlideDataExtension10);

            commonSlideData10.Append(shapeTree10);
            commonSlideData10.Append(commonSlideDataExtensionList10);

            ColorMapOverride colorMapOverride9 = new ColorMapOverride();
            A.MasterColorMapping masterColorMapping9 = new A.MasterColorMapping();

            colorMapOverride9.Append(masterColorMapping9);

            slideLayout8.Append(commonSlideData10);
            slideLayout8.Append(colorMapOverride9);

            slideLayoutPart8.SlideLayout = slideLayout8;
        }

        // Generates content of slideLayoutPart9.
        private void GenerateSlideLayoutPart9Content(SlideLayoutPart slideLayoutPart9)
        {
            SlideLayout slideLayout9 = new SlideLayout(){ Type = SlideLayoutValues.VerticalText, Preserve = true };
            slideLayout9.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            slideLayout9.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            slideLayout9.AddNamespaceDeclaration("p", "http://schemas.openxmlformats.org/presentationml/2006/main");

            CommonSlideData commonSlideData11 = new CommonSlideData(){ Name = "Nadpis a svislý text" };

            ShapeTree shapeTree11 = new ShapeTree();

            NonVisualGroupShapeProperties nonVisualGroupShapeProperties11 = new NonVisualGroupShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties59 = new NonVisualDrawingProperties(){ Id = (UInt32Value)1U, Name = "" };
            NonVisualGroupShapeDrawingProperties nonVisualGroupShapeDrawingProperties11 = new NonVisualGroupShapeDrawingProperties();
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties59 = new ApplicationNonVisualDrawingProperties();

            nonVisualGroupShapeProperties11.Append(nonVisualDrawingProperties59);
            nonVisualGroupShapeProperties11.Append(nonVisualGroupShapeDrawingProperties11);
            nonVisualGroupShapeProperties11.Append(applicationNonVisualDrawingProperties59);

            GroupShapeProperties groupShapeProperties11 = new GroupShapeProperties();

            A.TransformGroup transformGroup11 = new A.TransformGroup();
            A.Offset offset32 = new A.Offset(){ X = 0L, Y = 0L };
            A.Extents extents32 = new A.Extents(){ Cx = 0L, Cy = 0L };
            A.ChildOffset childOffset11 = new A.ChildOffset(){ X = 0L, Y = 0L };
            A.ChildExtents childExtents11 = new A.ChildExtents(){ Cx = 0L, Cy = 0L };

            transformGroup11.Append(offset32);
            transformGroup11.Append(extents32);
            transformGroup11.Append(childOffset11);
            transformGroup11.Append(childExtents11);

            groupShapeProperties11.Append(transformGroup11);

            Shape shape47 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties47 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties60 = new NonVisualDrawingProperties(){ Id = (UInt32Value)2U, Name = "Title 1" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties47 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks47 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties47.Append(shapeLocks47);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties60 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape47 = new PlaceholderShape(){ Type = PlaceholderValues.Title };

            applicationNonVisualDrawingProperties60.Append(placeholderShape47);

            nonVisualShapeProperties47.Append(nonVisualDrawingProperties60);
            nonVisualShapeProperties47.Append(nonVisualShapeDrawingProperties47);
            nonVisualShapeProperties47.Append(applicationNonVisualDrawingProperties60);
            ShapeProperties shapeProperties48 = new ShapeProperties();

            TextBody textBody47 = new TextBody();
            A.BodyProperties bodyProperties47 = new A.BodyProperties();
            A.ListStyle listStyle47 = new A.ListStyle();

            A.Paragraph paragraph71 = new A.Paragraph();

            A.Run run44 = new A.Run();
            A.RunProperties runProperties62 = new A.RunProperties(){ Language = "cs-CZ" };
            A.Text text62 = new A.Text();
            text62.Text = "Kliknutím lze upravit styl.";

            run44.Append(runProperties62);
            run44.Append(text62);
            A.EndParagraphRunProperties endParagraphRunProperties43 = new A.EndParagraphRunProperties(){ Language = "en-US", Dirty = false };

            paragraph71.Append(run44);
            paragraph71.Append(endParagraphRunProperties43);

            textBody47.Append(bodyProperties47);
            textBody47.Append(listStyle47);
            textBody47.Append(paragraph71);

            shape47.Append(nonVisualShapeProperties47);
            shape47.Append(shapeProperties48);
            shape47.Append(textBody47);

            Shape shape48 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties48 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties61 = new NonVisualDrawingProperties(){ Id = (UInt32Value)3U, Name = "Vertical Text Placeholder 2" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties48 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks48 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties48.Append(shapeLocks48);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties61 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape48 = new PlaceholderShape(){ Type = PlaceholderValues.Body, Orientation = DirectionValues.Vertical, Index = (UInt32Value)1U };

            applicationNonVisualDrawingProperties61.Append(placeholderShape48);

            nonVisualShapeProperties48.Append(nonVisualDrawingProperties61);
            nonVisualShapeProperties48.Append(nonVisualShapeDrawingProperties48);
            nonVisualShapeProperties48.Append(applicationNonVisualDrawingProperties61);
            ShapeProperties shapeProperties49 = new ShapeProperties();

            TextBody textBody48 = new TextBody();
            A.BodyProperties bodyProperties48 = new A.BodyProperties(){ Vertical = A.TextVerticalValues.EastAsianVetical };
            A.ListStyle listStyle48 = new A.ListStyle();

            A.Paragraph paragraph72 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties35 = new A.ParagraphProperties(){ Level = 0 };

            A.Run run45 = new A.Run();
            A.RunProperties runProperties63 = new A.RunProperties(){ Language = "cs-CZ" };
            A.Text text63 = new A.Text();
            text63.Text = "Po kliknutí můžete upravovat styly textu v předloze.";

            run45.Append(runProperties63);
            run45.Append(text63);

            paragraph72.Append(paragraphProperties35);
            paragraph72.Append(run45);

            A.Paragraph paragraph73 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties36 = new A.ParagraphProperties(){ Level = 1 };

            A.Run run46 = new A.Run();
            A.RunProperties runProperties64 = new A.RunProperties(){ Language = "cs-CZ" };
            A.Text text64 = new A.Text();
            text64.Text = "Druhá úroveň";

            run46.Append(runProperties64);
            run46.Append(text64);

            paragraph73.Append(paragraphProperties36);
            paragraph73.Append(run46);

            A.Paragraph paragraph74 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties37 = new A.ParagraphProperties(){ Level = 2 };

            A.Run run47 = new A.Run();
            A.RunProperties runProperties65 = new A.RunProperties(){ Language = "cs-CZ" };
            A.Text text65 = new A.Text();
            text65.Text = "Třetí úroveň";

            run47.Append(runProperties65);
            run47.Append(text65);

            paragraph74.Append(paragraphProperties37);
            paragraph74.Append(run47);

            A.Paragraph paragraph75 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties38 = new A.ParagraphProperties(){ Level = 3 };

            A.Run run48 = new A.Run();
            A.RunProperties runProperties66 = new A.RunProperties(){ Language = "cs-CZ" };
            A.Text text66 = new A.Text();
            text66.Text = "Čtvrtá úroveň";

            run48.Append(runProperties66);
            run48.Append(text66);

            paragraph75.Append(paragraphProperties38);
            paragraph75.Append(run48);

            A.Paragraph paragraph76 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties39 = new A.ParagraphProperties(){ Level = 4 };

            A.Run run49 = new A.Run();
            A.RunProperties runProperties67 = new A.RunProperties(){ Language = "cs-CZ" };
            A.Text text67 = new A.Text();
            text67.Text = "Pátá úroveň";

            run49.Append(runProperties67);
            run49.Append(text67);
            A.EndParagraphRunProperties endParagraphRunProperties44 = new A.EndParagraphRunProperties(){ Language = "en-US", Dirty = false };

            paragraph76.Append(paragraphProperties39);
            paragraph76.Append(run49);
            paragraph76.Append(endParagraphRunProperties44);

            textBody48.Append(bodyProperties48);
            textBody48.Append(listStyle48);
            textBody48.Append(paragraph72);
            textBody48.Append(paragraph73);
            textBody48.Append(paragraph74);
            textBody48.Append(paragraph75);
            textBody48.Append(paragraph76);

            shape48.Append(nonVisualShapeProperties48);
            shape48.Append(shapeProperties49);
            shape48.Append(textBody48);

            Shape shape49 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties49 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties62 = new NonVisualDrawingProperties(){ Id = (UInt32Value)4U, Name = "Date Placeholder 3" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties49 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks49 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties49.Append(shapeLocks49);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties62 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape49 = new PlaceholderShape(){ Type = PlaceholderValues.DateAndTime, Size = PlaceholderSizeValues.Half, Index = (UInt32Value)10U };

            applicationNonVisualDrawingProperties62.Append(placeholderShape49);

            nonVisualShapeProperties49.Append(nonVisualDrawingProperties62);
            nonVisualShapeProperties49.Append(nonVisualShapeDrawingProperties49);
            nonVisualShapeProperties49.Append(applicationNonVisualDrawingProperties62);
            ShapeProperties shapeProperties50 = new ShapeProperties();

            TextBody textBody49 = new TextBody();
            A.BodyProperties bodyProperties49 = new A.BodyProperties();
            A.ListStyle listStyle49 = new A.ListStyle();

            A.Paragraph paragraph77 = new A.Paragraph();

            A.Field field19 = new A.Field(){ Id = "{1D065981-0097-4BA9-A692-9FF652AFAF35}", Type = "datetimeFigureOut" };

            A.RunProperties runProperties68 = new A.RunProperties(){ Language = "cs-CZ" };
            runProperties68.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            A.Text text68 = new A.Text();
            text68.Text = "14.03.2023";

            field19.Append(runProperties68);
            field19.Append(text68);
            A.EndParagraphRunProperties endParagraphRunProperties45 = new A.EndParagraphRunProperties(){ Language = "cs-CZ" };

            paragraph77.Append(field19);
            paragraph77.Append(endParagraphRunProperties45);

            textBody49.Append(bodyProperties49);
            textBody49.Append(listStyle49);
            textBody49.Append(paragraph77);

            shape49.Append(nonVisualShapeProperties49);
            shape49.Append(shapeProperties50);
            shape49.Append(textBody49);

            Shape shape50 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties50 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties63 = new NonVisualDrawingProperties(){ Id = (UInt32Value)5U, Name = "Footer Placeholder 4" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties50 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks50 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties50.Append(shapeLocks50);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties63 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape50 = new PlaceholderShape(){ Type = PlaceholderValues.Footer, Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)11U };

            applicationNonVisualDrawingProperties63.Append(placeholderShape50);

            nonVisualShapeProperties50.Append(nonVisualDrawingProperties63);
            nonVisualShapeProperties50.Append(nonVisualShapeDrawingProperties50);
            nonVisualShapeProperties50.Append(applicationNonVisualDrawingProperties63);
            ShapeProperties shapeProperties51 = new ShapeProperties();

            TextBody textBody50 = new TextBody();
            A.BodyProperties bodyProperties50 = new A.BodyProperties();
            A.ListStyle listStyle50 = new A.ListStyle();

            A.Paragraph paragraph78 = new A.Paragraph();
            A.EndParagraphRunProperties endParagraphRunProperties46 = new A.EndParagraphRunProperties(){ Language = "cs-CZ" };

            paragraph78.Append(endParagraphRunProperties46);

            textBody50.Append(bodyProperties50);
            textBody50.Append(listStyle50);
            textBody50.Append(paragraph78);

            shape50.Append(nonVisualShapeProperties50);
            shape50.Append(shapeProperties51);
            shape50.Append(textBody50);

            Shape shape51 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties51 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties64 = new NonVisualDrawingProperties(){ Id = (UInt32Value)6U, Name = "Slide Number Placeholder 5" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties51 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks51 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties51.Append(shapeLocks51);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties64 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape51 = new PlaceholderShape(){ Type = PlaceholderValues.SlideNumber, Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)12U };

            applicationNonVisualDrawingProperties64.Append(placeholderShape51);

            nonVisualShapeProperties51.Append(nonVisualDrawingProperties64);
            nonVisualShapeProperties51.Append(nonVisualShapeDrawingProperties51);
            nonVisualShapeProperties51.Append(applicationNonVisualDrawingProperties64);
            ShapeProperties shapeProperties52 = new ShapeProperties();

            TextBody textBody51 = new TextBody();
            A.BodyProperties bodyProperties51 = new A.BodyProperties();
            A.ListStyle listStyle51 = new A.ListStyle();

            A.Paragraph paragraph79 = new A.Paragraph();

            A.Field field20 = new A.Field(){ Id = "{5D7F2B33-3E35-4F97-A448-9FEB151DF743}", Type = "slidenum" };

            A.RunProperties runProperties69 = new A.RunProperties(){ Language = "cs-CZ" };
            runProperties69.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            A.Text text69 = new A.Text();
            text69.Text = "‹#›";

            field20.Append(runProperties69);
            field20.Append(text69);
            A.EndParagraphRunProperties endParagraphRunProperties47 = new A.EndParagraphRunProperties(){ Language = "cs-CZ" };

            paragraph79.Append(field20);
            paragraph79.Append(endParagraphRunProperties47);

            textBody51.Append(bodyProperties51);
            textBody51.Append(listStyle51);
            textBody51.Append(paragraph79);

            shape51.Append(nonVisualShapeProperties51);
            shape51.Append(shapeProperties52);
            shape51.Append(textBody51);

            shapeTree11.Append(nonVisualGroupShapeProperties11);
            shapeTree11.Append(groupShapeProperties11);
            shapeTree11.Append(shape47);
            shapeTree11.Append(shape48);
            shapeTree11.Append(shape49);
            shapeTree11.Append(shape50);
            shapeTree11.Append(shape51);

            CommonSlideDataExtensionList commonSlideDataExtensionList11 = new CommonSlideDataExtensionList();

            CommonSlideDataExtension commonSlideDataExtension11 = new CommonSlideDataExtension(){ Uri = "{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}" };

            P14.CreationId creationId11 = new P14.CreationId(){ Val = (UInt32Value)248461432U };
            creationId11.AddNamespaceDeclaration("p14", "http://schemas.microsoft.com/office/powerpoint/2010/main");

            commonSlideDataExtension11.Append(creationId11);

            commonSlideDataExtensionList11.Append(commonSlideDataExtension11);

            commonSlideData11.Append(shapeTree11);
            commonSlideData11.Append(commonSlideDataExtensionList11);

            ColorMapOverride colorMapOverride10 = new ColorMapOverride();
            A.MasterColorMapping masterColorMapping10 = new A.MasterColorMapping();

            colorMapOverride10.Append(masterColorMapping10);

            slideLayout9.Append(commonSlideData11);
            slideLayout9.Append(colorMapOverride10);

            slideLayoutPart9.SlideLayout = slideLayout9;
        }

        // Generates content of slideLayoutPart10.
        private void GenerateSlideLayoutPart10Content(SlideLayoutPart slideLayoutPart10)
        {
            SlideLayout slideLayout10 = new SlideLayout(){ Type = SlideLayoutValues.TwoObjects, Preserve = true };
            slideLayout10.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            slideLayout10.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            slideLayout10.AddNamespaceDeclaration("p", "http://schemas.openxmlformats.org/presentationml/2006/main");

            CommonSlideData commonSlideData12 = new CommonSlideData(){ Name = "Dva obsahy" };

            ShapeTree shapeTree12 = new ShapeTree();

            NonVisualGroupShapeProperties nonVisualGroupShapeProperties12 = new NonVisualGroupShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties65 = new NonVisualDrawingProperties(){ Id = (UInt32Value)1U, Name = "" };
            NonVisualGroupShapeDrawingProperties nonVisualGroupShapeDrawingProperties12 = new NonVisualGroupShapeDrawingProperties();
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties65 = new ApplicationNonVisualDrawingProperties();

            nonVisualGroupShapeProperties12.Append(nonVisualDrawingProperties65);
            nonVisualGroupShapeProperties12.Append(nonVisualGroupShapeDrawingProperties12);
            nonVisualGroupShapeProperties12.Append(applicationNonVisualDrawingProperties65);

            GroupShapeProperties groupShapeProperties12 = new GroupShapeProperties();

            A.TransformGroup transformGroup12 = new A.TransformGroup();
            A.Offset offset33 = new A.Offset(){ X = 0L, Y = 0L };
            A.Extents extents33 = new A.Extents(){ Cx = 0L, Cy = 0L };
            A.ChildOffset childOffset12 = new A.ChildOffset(){ X = 0L, Y = 0L };
            A.ChildExtents childExtents12 = new A.ChildExtents(){ Cx = 0L, Cy = 0L };

            transformGroup12.Append(offset33);
            transformGroup12.Append(extents33);
            transformGroup12.Append(childOffset12);
            transformGroup12.Append(childExtents12);

            groupShapeProperties12.Append(transformGroup12);

            Shape shape52 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties52 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties66 = new NonVisualDrawingProperties(){ Id = (UInt32Value)2U, Name = "Title 1" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties52 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks52 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties52.Append(shapeLocks52);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties66 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape52 = new PlaceholderShape(){ Type = PlaceholderValues.Title };

            applicationNonVisualDrawingProperties66.Append(placeholderShape52);

            nonVisualShapeProperties52.Append(nonVisualDrawingProperties66);
            nonVisualShapeProperties52.Append(nonVisualShapeDrawingProperties52);
            nonVisualShapeProperties52.Append(applicationNonVisualDrawingProperties66);
            ShapeProperties shapeProperties53 = new ShapeProperties();

            TextBody textBody52 = new TextBody();
            A.BodyProperties bodyProperties52 = new A.BodyProperties();
            A.ListStyle listStyle52 = new A.ListStyle();

            A.Paragraph paragraph80 = new A.Paragraph();

            A.Run run50 = new A.Run();
            A.RunProperties runProperties70 = new A.RunProperties(){ Language = "cs-CZ" };
            A.Text text70 = new A.Text();
            text70.Text = "Kliknutím lze upravit styl.";

            run50.Append(runProperties70);
            run50.Append(text70);
            A.EndParagraphRunProperties endParagraphRunProperties48 = new A.EndParagraphRunProperties(){ Language = "en-US", Dirty = false };

            paragraph80.Append(run50);
            paragraph80.Append(endParagraphRunProperties48);

            textBody52.Append(bodyProperties52);
            textBody52.Append(listStyle52);
            textBody52.Append(paragraph80);

            shape52.Append(nonVisualShapeProperties52);
            shape52.Append(shapeProperties53);
            shape52.Append(textBody52);

            Shape shape53 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties53 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties67 = new NonVisualDrawingProperties(){ Id = (UInt32Value)3U, Name = "Content Placeholder 2" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties53 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks53 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties53.Append(shapeLocks53);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties67 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape53 = new PlaceholderShape(){ Size = PlaceholderSizeValues.Half, Index = (UInt32Value)1U };

            applicationNonVisualDrawingProperties67.Append(placeholderShape53);

            nonVisualShapeProperties53.Append(nonVisualDrawingProperties67);
            nonVisualShapeProperties53.Append(nonVisualShapeDrawingProperties53);
            nonVisualShapeProperties53.Append(applicationNonVisualDrawingProperties67);

            ShapeProperties shapeProperties54 = new ShapeProperties();

            A.Transform2D transform2D21 = new A.Transform2D();
            A.Offset offset34 = new A.Offset(){ X = 628650L, Y = 1825625L };
            A.Extents extents34 = new A.Extents(){ Cx = 3886200L, Cy = 4351338L };

            transform2D21.Append(offset34);
            transform2D21.Append(extents34);

            shapeProperties54.Append(transform2D21);

            TextBody textBody53 = new TextBody();
            A.BodyProperties bodyProperties53 = new A.BodyProperties();
            A.ListStyle listStyle53 = new A.ListStyle();

            A.Paragraph paragraph81 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties40 = new A.ParagraphProperties(){ Level = 0 };

            A.Run run51 = new A.Run();
            A.RunProperties runProperties71 = new A.RunProperties(){ Language = "cs-CZ" };
            A.Text text71 = new A.Text();
            text71.Text = "Po kliknutí můžete upravovat styly textu v předloze.";

            run51.Append(runProperties71);
            run51.Append(text71);

            paragraph81.Append(paragraphProperties40);
            paragraph81.Append(run51);

            A.Paragraph paragraph82 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties41 = new A.ParagraphProperties(){ Level = 1 };

            A.Run run52 = new A.Run();
            A.RunProperties runProperties72 = new A.RunProperties(){ Language = "cs-CZ" };
            A.Text text72 = new A.Text();
            text72.Text = "Druhá úroveň";

            run52.Append(runProperties72);
            run52.Append(text72);

            paragraph82.Append(paragraphProperties41);
            paragraph82.Append(run52);

            A.Paragraph paragraph83 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties42 = new A.ParagraphProperties(){ Level = 2 };

            A.Run run53 = new A.Run();
            A.RunProperties runProperties73 = new A.RunProperties(){ Language = "cs-CZ" };
            A.Text text73 = new A.Text();
            text73.Text = "Třetí úroveň";

            run53.Append(runProperties73);
            run53.Append(text73);

            paragraph83.Append(paragraphProperties42);
            paragraph83.Append(run53);

            A.Paragraph paragraph84 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties43 = new A.ParagraphProperties(){ Level = 3 };

            A.Run run54 = new A.Run();
            A.RunProperties runProperties74 = new A.RunProperties(){ Language = "cs-CZ" };
            A.Text text74 = new A.Text();
            text74.Text = "Čtvrtá úroveň";

            run54.Append(runProperties74);
            run54.Append(text74);

            paragraph84.Append(paragraphProperties43);
            paragraph84.Append(run54);

            A.Paragraph paragraph85 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties44 = new A.ParagraphProperties(){ Level = 4 };

            A.Run run55 = new A.Run();
            A.RunProperties runProperties75 = new A.RunProperties(){ Language = "cs-CZ" };
            A.Text text75 = new A.Text();
            text75.Text = "Pátá úroveň";

            run55.Append(runProperties75);
            run55.Append(text75);
            A.EndParagraphRunProperties endParagraphRunProperties49 = new A.EndParagraphRunProperties(){ Language = "en-US", Dirty = false };

            paragraph85.Append(paragraphProperties44);
            paragraph85.Append(run55);
            paragraph85.Append(endParagraphRunProperties49);

            textBody53.Append(bodyProperties53);
            textBody53.Append(listStyle53);
            textBody53.Append(paragraph81);
            textBody53.Append(paragraph82);
            textBody53.Append(paragraph83);
            textBody53.Append(paragraph84);
            textBody53.Append(paragraph85);

            shape53.Append(nonVisualShapeProperties53);
            shape53.Append(shapeProperties54);
            shape53.Append(textBody53);

            Shape shape54 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties54 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties68 = new NonVisualDrawingProperties(){ Id = (UInt32Value)4U, Name = "Content Placeholder 3" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties54 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks54 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties54.Append(shapeLocks54);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties68 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape54 = new PlaceholderShape(){ Size = PlaceholderSizeValues.Half, Index = (UInt32Value)2U };

            applicationNonVisualDrawingProperties68.Append(placeholderShape54);

            nonVisualShapeProperties54.Append(nonVisualDrawingProperties68);
            nonVisualShapeProperties54.Append(nonVisualShapeDrawingProperties54);
            nonVisualShapeProperties54.Append(applicationNonVisualDrawingProperties68);

            ShapeProperties shapeProperties55 = new ShapeProperties();

            A.Transform2D transform2D22 = new A.Transform2D();
            A.Offset offset35 = new A.Offset(){ X = 4629150L, Y = 1825625L };
            A.Extents extents35 = new A.Extents(){ Cx = 3886200L, Cy = 4351338L };

            transform2D22.Append(offset35);
            transform2D22.Append(extents35);

            shapeProperties55.Append(transform2D22);

            TextBody textBody54 = new TextBody();
            A.BodyProperties bodyProperties54 = new A.BodyProperties();
            A.ListStyle listStyle54 = new A.ListStyle();

            A.Paragraph paragraph86 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties45 = new A.ParagraphProperties(){ Level = 0 };

            A.Run run56 = new A.Run();
            A.RunProperties runProperties76 = new A.RunProperties(){ Language = "cs-CZ" };
            A.Text text76 = new A.Text();
            text76.Text = "Po kliknutí můžete upravovat styly textu v předloze.";

            run56.Append(runProperties76);
            run56.Append(text76);

            paragraph86.Append(paragraphProperties45);
            paragraph86.Append(run56);

            A.Paragraph paragraph87 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties46 = new A.ParagraphProperties(){ Level = 1 };

            A.Run run57 = new A.Run();
            A.RunProperties runProperties77 = new A.RunProperties(){ Language = "cs-CZ" };
            A.Text text77 = new A.Text();
            text77.Text = "Druhá úroveň";

            run57.Append(runProperties77);
            run57.Append(text77);

            paragraph87.Append(paragraphProperties46);
            paragraph87.Append(run57);

            A.Paragraph paragraph88 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties47 = new A.ParagraphProperties(){ Level = 2 };

            A.Run run58 = new A.Run();
            A.RunProperties runProperties78 = new A.RunProperties(){ Language = "cs-CZ" };
            A.Text text78 = new A.Text();
            text78.Text = "Třetí úroveň";

            run58.Append(runProperties78);
            run58.Append(text78);

            paragraph88.Append(paragraphProperties47);
            paragraph88.Append(run58);

            A.Paragraph paragraph89 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties48 = new A.ParagraphProperties(){ Level = 3 };

            A.Run run59 = new A.Run();
            A.RunProperties runProperties79 = new A.RunProperties(){ Language = "cs-CZ" };
            A.Text text79 = new A.Text();
            text79.Text = "Čtvrtá úroveň";

            run59.Append(runProperties79);
            run59.Append(text79);

            paragraph89.Append(paragraphProperties48);
            paragraph89.Append(run59);

            A.Paragraph paragraph90 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties49 = new A.ParagraphProperties(){ Level = 4 };

            A.Run run60 = new A.Run();
            A.RunProperties runProperties80 = new A.RunProperties(){ Language = "cs-CZ" };
            A.Text text80 = new A.Text();
            text80.Text = "Pátá úroveň";

            run60.Append(runProperties80);
            run60.Append(text80);
            A.EndParagraphRunProperties endParagraphRunProperties50 = new A.EndParagraphRunProperties(){ Language = "en-US", Dirty = false };

            paragraph90.Append(paragraphProperties49);
            paragraph90.Append(run60);
            paragraph90.Append(endParagraphRunProperties50);

            textBody54.Append(bodyProperties54);
            textBody54.Append(listStyle54);
            textBody54.Append(paragraph86);
            textBody54.Append(paragraph87);
            textBody54.Append(paragraph88);
            textBody54.Append(paragraph89);
            textBody54.Append(paragraph90);

            shape54.Append(nonVisualShapeProperties54);
            shape54.Append(shapeProperties55);
            shape54.Append(textBody54);

            Shape shape55 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties55 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties69 = new NonVisualDrawingProperties(){ Id = (UInt32Value)5U, Name = "Date Placeholder 4" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties55 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks55 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties55.Append(shapeLocks55);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties69 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape55 = new PlaceholderShape(){ Type = PlaceholderValues.DateAndTime, Size = PlaceholderSizeValues.Half, Index = (UInt32Value)10U };

            applicationNonVisualDrawingProperties69.Append(placeholderShape55);

            nonVisualShapeProperties55.Append(nonVisualDrawingProperties69);
            nonVisualShapeProperties55.Append(nonVisualShapeDrawingProperties55);
            nonVisualShapeProperties55.Append(applicationNonVisualDrawingProperties69);
            ShapeProperties shapeProperties56 = new ShapeProperties();

            TextBody textBody55 = new TextBody();
            A.BodyProperties bodyProperties55 = new A.BodyProperties();
            A.ListStyle listStyle55 = new A.ListStyle();

            A.Paragraph paragraph91 = new A.Paragraph();

            A.Field field21 = new A.Field(){ Id = "{1D065981-0097-4BA9-A692-9FF652AFAF35}", Type = "datetimeFigureOut" };

            A.RunProperties runProperties81 = new A.RunProperties(){ Language = "cs-CZ" };
            runProperties81.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            A.Text text81 = new A.Text();
            text81.Text = "14.03.2023";

            field21.Append(runProperties81);
            field21.Append(text81);
            A.EndParagraphRunProperties endParagraphRunProperties51 = new A.EndParagraphRunProperties(){ Language = "cs-CZ" };

            paragraph91.Append(field21);
            paragraph91.Append(endParagraphRunProperties51);

            textBody55.Append(bodyProperties55);
            textBody55.Append(listStyle55);
            textBody55.Append(paragraph91);

            shape55.Append(nonVisualShapeProperties55);
            shape55.Append(shapeProperties56);
            shape55.Append(textBody55);

            Shape shape56 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties56 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties70 = new NonVisualDrawingProperties(){ Id = (UInt32Value)6U, Name = "Footer Placeholder 5" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties56 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks56 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties56.Append(shapeLocks56);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties70 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape56 = new PlaceholderShape(){ Type = PlaceholderValues.Footer, Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)11U };

            applicationNonVisualDrawingProperties70.Append(placeholderShape56);

            nonVisualShapeProperties56.Append(nonVisualDrawingProperties70);
            nonVisualShapeProperties56.Append(nonVisualShapeDrawingProperties56);
            nonVisualShapeProperties56.Append(applicationNonVisualDrawingProperties70);
            ShapeProperties shapeProperties57 = new ShapeProperties();

            TextBody textBody56 = new TextBody();
            A.BodyProperties bodyProperties56 = new A.BodyProperties();
            A.ListStyle listStyle56 = new A.ListStyle();

            A.Paragraph paragraph92 = new A.Paragraph();
            A.EndParagraphRunProperties endParagraphRunProperties52 = new A.EndParagraphRunProperties(){ Language = "cs-CZ" };

            paragraph92.Append(endParagraphRunProperties52);

            textBody56.Append(bodyProperties56);
            textBody56.Append(listStyle56);
            textBody56.Append(paragraph92);

            shape56.Append(nonVisualShapeProperties56);
            shape56.Append(shapeProperties57);
            shape56.Append(textBody56);

            Shape shape57 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties57 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties71 = new NonVisualDrawingProperties(){ Id = (UInt32Value)7U, Name = "Slide Number Placeholder 6" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties57 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks57 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties57.Append(shapeLocks57);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties71 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape57 = new PlaceholderShape(){ Type = PlaceholderValues.SlideNumber, Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)12U };

            applicationNonVisualDrawingProperties71.Append(placeholderShape57);

            nonVisualShapeProperties57.Append(nonVisualDrawingProperties71);
            nonVisualShapeProperties57.Append(nonVisualShapeDrawingProperties57);
            nonVisualShapeProperties57.Append(applicationNonVisualDrawingProperties71);
            ShapeProperties shapeProperties58 = new ShapeProperties();

            TextBody textBody57 = new TextBody();
            A.BodyProperties bodyProperties57 = new A.BodyProperties();
            A.ListStyle listStyle57 = new A.ListStyle();

            A.Paragraph paragraph93 = new A.Paragraph();

            A.Field field22 = new A.Field(){ Id = "{5D7F2B33-3E35-4F97-A448-9FEB151DF743}", Type = "slidenum" };

            A.RunProperties runProperties82 = new A.RunProperties(){ Language = "cs-CZ" };
            runProperties82.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            A.Text text82 = new A.Text();
            text82.Text = "‹#›";

            field22.Append(runProperties82);
            field22.Append(text82);
            A.EndParagraphRunProperties endParagraphRunProperties53 = new A.EndParagraphRunProperties(){ Language = "cs-CZ" };

            paragraph93.Append(field22);
            paragraph93.Append(endParagraphRunProperties53);

            textBody57.Append(bodyProperties57);
            textBody57.Append(listStyle57);
            textBody57.Append(paragraph93);

            shape57.Append(nonVisualShapeProperties57);
            shape57.Append(shapeProperties58);
            shape57.Append(textBody57);

            shapeTree12.Append(nonVisualGroupShapeProperties12);
            shapeTree12.Append(groupShapeProperties12);
            shapeTree12.Append(shape52);
            shapeTree12.Append(shape53);
            shapeTree12.Append(shape54);
            shapeTree12.Append(shape55);
            shapeTree12.Append(shape56);
            shapeTree12.Append(shape57);

            CommonSlideDataExtensionList commonSlideDataExtensionList12 = new CommonSlideDataExtensionList();

            CommonSlideDataExtension commonSlideDataExtension12 = new CommonSlideDataExtension(){ Uri = "{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}" };

            P14.CreationId creationId12 = new P14.CreationId(){ Val = (UInt32Value)3521598879U };
            creationId12.AddNamespaceDeclaration("p14", "http://schemas.microsoft.com/office/powerpoint/2010/main");

            commonSlideDataExtension12.Append(creationId12);

            commonSlideDataExtensionList12.Append(commonSlideDataExtension12);

            commonSlideData12.Append(shapeTree12);
            commonSlideData12.Append(commonSlideDataExtensionList12);

            ColorMapOverride colorMapOverride11 = new ColorMapOverride();
            A.MasterColorMapping masterColorMapping11 = new A.MasterColorMapping();

            colorMapOverride11.Append(masterColorMapping11);

            slideLayout10.Append(commonSlideData12);
            slideLayout10.Append(colorMapOverride11);

            slideLayoutPart10.SlideLayout = slideLayout10;
        }

        // Generates content of slideLayoutPart11.
        private void GenerateSlideLayoutPart11Content(SlideLayoutPart slideLayoutPart11)
        {
            SlideLayout slideLayout11 = new SlideLayout(){ Type = SlideLayoutValues.PictureText, Preserve = true };
            slideLayout11.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            slideLayout11.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            slideLayout11.AddNamespaceDeclaration("p", "http://schemas.openxmlformats.org/presentationml/2006/main");

            CommonSlideData commonSlideData13 = new CommonSlideData(){ Name = "Obrázek s titulkem" };

            ShapeTree shapeTree13 = new ShapeTree();

            NonVisualGroupShapeProperties nonVisualGroupShapeProperties13 = new NonVisualGroupShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties72 = new NonVisualDrawingProperties(){ Id = (UInt32Value)1U, Name = "" };
            NonVisualGroupShapeDrawingProperties nonVisualGroupShapeDrawingProperties13 = new NonVisualGroupShapeDrawingProperties();
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties72 = new ApplicationNonVisualDrawingProperties();

            nonVisualGroupShapeProperties13.Append(nonVisualDrawingProperties72);
            nonVisualGroupShapeProperties13.Append(nonVisualGroupShapeDrawingProperties13);
            nonVisualGroupShapeProperties13.Append(applicationNonVisualDrawingProperties72);

            GroupShapeProperties groupShapeProperties13 = new GroupShapeProperties();

            A.TransformGroup transformGroup13 = new A.TransformGroup();
            A.Offset offset36 = new A.Offset(){ X = 0L, Y = 0L };
            A.Extents extents36 = new A.Extents(){ Cx = 0L, Cy = 0L };
            A.ChildOffset childOffset13 = new A.ChildOffset(){ X = 0L, Y = 0L };
            A.ChildExtents childExtents13 = new A.ChildExtents(){ Cx = 0L, Cy = 0L };

            transformGroup13.Append(offset36);
            transformGroup13.Append(extents36);
            transformGroup13.Append(childOffset13);
            transformGroup13.Append(childExtents13);

            groupShapeProperties13.Append(transformGroup13);

            Shape shape58 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties58 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties73 = new NonVisualDrawingProperties(){ Id = (UInt32Value)2U, Name = "Title 1" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties58 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks58 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties58.Append(shapeLocks58);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties73 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape58 = new PlaceholderShape(){ Type = PlaceholderValues.Title };

            applicationNonVisualDrawingProperties73.Append(placeholderShape58);

            nonVisualShapeProperties58.Append(nonVisualDrawingProperties73);
            nonVisualShapeProperties58.Append(nonVisualShapeDrawingProperties58);
            nonVisualShapeProperties58.Append(applicationNonVisualDrawingProperties73);

            ShapeProperties shapeProperties59 = new ShapeProperties();

            A.Transform2D transform2D23 = new A.Transform2D();
            A.Offset offset37 = new A.Offset(){ X = 629841L, Y = 457200L };
            A.Extents extents37 = new A.Extents(){ Cx = 2949178L, Cy = 1600200L };

            transform2D23.Append(offset37);
            transform2D23.Append(extents37);

            shapeProperties59.Append(transform2D23);

            TextBody textBody58 = new TextBody();
            A.BodyProperties bodyProperties58 = new A.BodyProperties(){ Anchor = A.TextAnchoringTypeValues.Bottom };

            A.ListStyle listStyle58 = new A.ListStyle();

            A.Level1ParagraphProperties level1ParagraphProperties17 = new A.Level1ParagraphProperties();
            A.DefaultRunProperties defaultRunProperties91 = new A.DefaultRunProperties(){ FontSize = 3200 };

            level1ParagraphProperties17.Append(defaultRunProperties91);

            listStyle58.Append(level1ParagraphProperties17);

            A.Paragraph paragraph94 = new A.Paragraph();

            A.Run run61 = new A.Run();
            A.RunProperties runProperties83 = new A.RunProperties(){ Language = "cs-CZ" };
            A.Text text83 = new A.Text();
            text83.Text = "Kliknutím lze upravit styl.";

            run61.Append(runProperties83);
            run61.Append(text83);
            A.EndParagraphRunProperties endParagraphRunProperties54 = new A.EndParagraphRunProperties(){ Language = "en-US", Dirty = false };

            paragraph94.Append(run61);
            paragraph94.Append(endParagraphRunProperties54);

            textBody58.Append(bodyProperties58);
            textBody58.Append(listStyle58);
            textBody58.Append(paragraph94);

            shape58.Append(nonVisualShapeProperties58);
            shape58.Append(shapeProperties59);
            shape58.Append(textBody58);

            Shape shape59 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties59 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties74 = new NonVisualDrawingProperties(){ Id = (UInt32Value)3U, Name = "Picture Placeholder 2" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties59 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks59 = new A.ShapeLocks(){ NoGrouping = true, NoChangeAspect = true };

            nonVisualShapeDrawingProperties59.Append(shapeLocks59);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties74 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape59 = new PlaceholderShape(){ Type = PlaceholderValues.Picture, Index = (UInt32Value)1U };

            applicationNonVisualDrawingProperties74.Append(placeholderShape59);

            nonVisualShapeProperties59.Append(nonVisualDrawingProperties74);
            nonVisualShapeProperties59.Append(nonVisualShapeDrawingProperties59);
            nonVisualShapeProperties59.Append(applicationNonVisualDrawingProperties74);

            ShapeProperties shapeProperties60 = new ShapeProperties();

            A.Transform2D transform2D24 = new A.Transform2D();
            A.Offset offset38 = new A.Offset(){ X = 3887391L, Y = 987426L };
            A.Extents extents38 = new A.Extents(){ Cx = 4629150L, Cy = 4873625L };

            transform2D24.Append(offset38);
            transform2D24.Append(extents38);

            shapeProperties60.Append(transform2D24);

            TextBody textBody59 = new TextBody();
            A.BodyProperties bodyProperties59 = new A.BodyProperties(){ Anchor = A.TextAnchoringTypeValues.Top };

            A.ListStyle listStyle59 = new A.ListStyle();

            A.Level1ParagraphProperties level1ParagraphProperties18 = new A.Level1ParagraphProperties(){ LeftMargin = 0, Indent = 0 };
            A.NoBullet noBullet47 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties92 = new A.DefaultRunProperties(){ FontSize = 3200 };

            level1ParagraphProperties18.Append(noBullet47);
            level1ParagraphProperties18.Append(defaultRunProperties92);

            A.Level2ParagraphProperties level2ParagraphProperties10 = new A.Level2ParagraphProperties(){ LeftMargin = 457200, Indent = 0 };
            A.NoBullet noBullet48 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties93 = new A.DefaultRunProperties(){ FontSize = 2800 };

            level2ParagraphProperties10.Append(noBullet48);
            level2ParagraphProperties10.Append(defaultRunProperties93);

            A.Level3ParagraphProperties level3ParagraphProperties10 = new A.Level3ParagraphProperties(){ LeftMargin = 914400, Indent = 0 };
            A.NoBullet noBullet49 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties94 = new A.DefaultRunProperties(){ FontSize = 2400 };

            level3ParagraphProperties10.Append(noBullet49);
            level3ParagraphProperties10.Append(defaultRunProperties94);

            A.Level4ParagraphProperties level4ParagraphProperties10 = new A.Level4ParagraphProperties(){ LeftMargin = 1371600, Indent = 0 };
            A.NoBullet noBullet50 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties95 = new A.DefaultRunProperties(){ FontSize = 2000 };

            level4ParagraphProperties10.Append(noBullet50);
            level4ParagraphProperties10.Append(defaultRunProperties95);

            A.Level5ParagraphProperties level5ParagraphProperties10 = new A.Level5ParagraphProperties(){ LeftMargin = 1828800, Indent = 0 };
            A.NoBullet noBullet51 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties96 = new A.DefaultRunProperties(){ FontSize = 2000 };

            level5ParagraphProperties10.Append(noBullet51);
            level5ParagraphProperties10.Append(defaultRunProperties96);

            A.Level6ParagraphProperties level6ParagraphProperties10 = new A.Level6ParagraphProperties(){ LeftMargin = 2286000, Indent = 0 };
            A.NoBullet noBullet52 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties97 = new A.DefaultRunProperties(){ FontSize = 2000 };

            level6ParagraphProperties10.Append(noBullet52);
            level6ParagraphProperties10.Append(defaultRunProperties97);

            A.Level7ParagraphProperties level7ParagraphProperties10 = new A.Level7ParagraphProperties(){ LeftMargin = 2743200, Indent = 0 };
            A.NoBullet noBullet53 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties98 = new A.DefaultRunProperties(){ FontSize = 2000 };

            level7ParagraphProperties10.Append(noBullet53);
            level7ParagraphProperties10.Append(defaultRunProperties98);

            A.Level8ParagraphProperties level8ParagraphProperties10 = new A.Level8ParagraphProperties(){ LeftMargin = 3200400, Indent = 0 };
            A.NoBullet noBullet54 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties99 = new A.DefaultRunProperties(){ FontSize = 2000 };

            level8ParagraphProperties10.Append(noBullet54);
            level8ParagraphProperties10.Append(defaultRunProperties99);

            A.Level9ParagraphProperties level9ParagraphProperties10 = new A.Level9ParagraphProperties(){ LeftMargin = 3657600, Indent = 0 };
            A.NoBullet noBullet55 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties100 = new A.DefaultRunProperties(){ FontSize = 2000 };

            level9ParagraphProperties10.Append(noBullet55);
            level9ParagraphProperties10.Append(defaultRunProperties100);

            listStyle59.Append(level1ParagraphProperties18);
            listStyle59.Append(level2ParagraphProperties10);
            listStyle59.Append(level3ParagraphProperties10);
            listStyle59.Append(level4ParagraphProperties10);
            listStyle59.Append(level5ParagraphProperties10);
            listStyle59.Append(level6ParagraphProperties10);
            listStyle59.Append(level7ParagraphProperties10);
            listStyle59.Append(level8ParagraphProperties10);
            listStyle59.Append(level9ParagraphProperties10);

            A.Paragraph paragraph95 = new A.Paragraph();

            A.Run run62 = new A.Run();
            A.RunProperties runProperties84 = new A.RunProperties(){ Language = "cs-CZ" };
            A.Text text84 = new A.Text();
            text84.Text = "Kliknutím na ikonu přidáte obrázek.";

            run62.Append(runProperties84);
            run62.Append(text84);
            A.EndParagraphRunProperties endParagraphRunProperties55 = new A.EndParagraphRunProperties(){ Language = "en-US", Dirty = false };

            paragraph95.Append(run62);
            paragraph95.Append(endParagraphRunProperties55);

            textBody59.Append(bodyProperties59);
            textBody59.Append(listStyle59);
            textBody59.Append(paragraph95);

            shape59.Append(nonVisualShapeProperties59);
            shape59.Append(shapeProperties60);
            shape59.Append(textBody59);

            Shape shape60 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties60 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties75 = new NonVisualDrawingProperties(){ Id = (UInt32Value)4U, Name = "Text Placeholder 3" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties60 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks60 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties60.Append(shapeLocks60);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties75 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape60 = new PlaceholderShape(){ Type = PlaceholderValues.Body, Size = PlaceholderSizeValues.Half, Index = (UInt32Value)2U };

            applicationNonVisualDrawingProperties75.Append(placeholderShape60);

            nonVisualShapeProperties60.Append(nonVisualDrawingProperties75);
            nonVisualShapeProperties60.Append(nonVisualShapeDrawingProperties60);
            nonVisualShapeProperties60.Append(applicationNonVisualDrawingProperties75);

            ShapeProperties shapeProperties61 = new ShapeProperties();

            A.Transform2D transform2D25 = new A.Transform2D();
            A.Offset offset39 = new A.Offset(){ X = 629841L, Y = 2057400L };
            A.Extents extents39 = new A.Extents(){ Cx = 2949178L, Cy = 3811588L };

            transform2D25.Append(offset39);
            transform2D25.Append(extents39);

            shapeProperties61.Append(transform2D25);

            TextBody textBody60 = new TextBody();
            A.BodyProperties bodyProperties60 = new A.BodyProperties();

            A.ListStyle listStyle60 = new A.ListStyle();

            A.Level1ParagraphProperties level1ParagraphProperties19 = new A.Level1ParagraphProperties(){ LeftMargin = 0, Indent = 0 };
            A.NoBullet noBullet56 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties101 = new A.DefaultRunProperties(){ FontSize = 1600 };

            level1ParagraphProperties19.Append(noBullet56);
            level1ParagraphProperties19.Append(defaultRunProperties101);

            A.Level2ParagraphProperties level2ParagraphProperties11 = new A.Level2ParagraphProperties(){ LeftMargin = 457200, Indent = 0 };
            A.NoBullet noBullet57 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties102 = new A.DefaultRunProperties(){ FontSize = 1400 };

            level2ParagraphProperties11.Append(noBullet57);
            level2ParagraphProperties11.Append(defaultRunProperties102);

            A.Level3ParagraphProperties level3ParagraphProperties11 = new A.Level3ParagraphProperties(){ LeftMargin = 914400, Indent = 0 };
            A.NoBullet noBullet58 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties103 = new A.DefaultRunProperties(){ FontSize = 1200 };

            level3ParagraphProperties11.Append(noBullet58);
            level3ParagraphProperties11.Append(defaultRunProperties103);

            A.Level4ParagraphProperties level4ParagraphProperties11 = new A.Level4ParagraphProperties(){ LeftMargin = 1371600, Indent = 0 };
            A.NoBullet noBullet59 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties104 = new A.DefaultRunProperties(){ FontSize = 1000 };

            level4ParagraphProperties11.Append(noBullet59);
            level4ParagraphProperties11.Append(defaultRunProperties104);

            A.Level5ParagraphProperties level5ParagraphProperties11 = new A.Level5ParagraphProperties(){ LeftMargin = 1828800, Indent = 0 };
            A.NoBullet noBullet60 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties105 = new A.DefaultRunProperties(){ FontSize = 1000 };

            level5ParagraphProperties11.Append(noBullet60);
            level5ParagraphProperties11.Append(defaultRunProperties105);

            A.Level6ParagraphProperties level6ParagraphProperties11 = new A.Level6ParagraphProperties(){ LeftMargin = 2286000, Indent = 0 };
            A.NoBullet noBullet61 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties106 = new A.DefaultRunProperties(){ FontSize = 1000 };

            level6ParagraphProperties11.Append(noBullet61);
            level6ParagraphProperties11.Append(defaultRunProperties106);

            A.Level7ParagraphProperties level7ParagraphProperties11 = new A.Level7ParagraphProperties(){ LeftMargin = 2743200, Indent = 0 };
            A.NoBullet noBullet62 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties107 = new A.DefaultRunProperties(){ FontSize = 1000 };

            level7ParagraphProperties11.Append(noBullet62);
            level7ParagraphProperties11.Append(defaultRunProperties107);

            A.Level8ParagraphProperties level8ParagraphProperties11 = new A.Level8ParagraphProperties(){ LeftMargin = 3200400, Indent = 0 };
            A.NoBullet noBullet63 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties108 = new A.DefaultRunProperties(){ FontSize = 1000 };

            level8ParagraphProperties11.Append(noBullet63);
            level8ParagraphProperties11.Append(defaultRunProperties108);

            A.Level9ParagraphProperties level9ParagraphProperties11 = new A.Level9ParagraphProperties(){ LeftMargin = 3657600, Indent = 0 };
            A.NoBullet noBullet64 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties109 = new A.DefaultRunProperties(){ FontSize = 1000 };

            level9ParagraphProperties11.Append(noBullet64);
            level9ParagraphProperties11.Append(defaultRunProperties109);

            listStyle60.Append(level1ParagraphProperties19);
            listStyle60.Append(level2ParagraphProperties11);
            listStyle60.Append(level3ParagraphProperties11);
            listStyle60.Append(level4ParagraphProperties11);
            listStyle60.Append(level5ParagraphProperties11);
            listStyle60.Append(level6ParagraphProperties11);
            listStyle60.Append(level7ParagraphProperties11);
            listStyle60.Append(level8ParagraphProperties11);
            listStyle60.Append(level9ParagraphProperties11);

            A.Paragraph paragraph96 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties50 = new A.ParagraphProperties(){ Level = 0 };

            A.Run run63 = new A.Run();
            A.RunProperties runProperties85 = new A.RunProperties(){ Language = "cs-CZ" };
            A.Text text85 = new A.Text();
            text85.Text = "Po kliknutí můžete upravovat styly textu v předloze.";

            run63.Append(runProperties85);
            run63.Append(text85);

            paragraph96.Append(paragraphProperties50);
            paragraph96.Append(run63);

            textBody60.Append(bodyProperties60);
            textBody60.Append(listStyle60);
            textBody60.Append(paragraph96);

            shape60.Append(nonVisualShapeProperties60);
            shape60.Append(shapeProperties61);
            shape60.Append(textBody60);

            Shape shape61 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties61 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties76 = new NonVisualDrawingProperties(){ Id = (UInt32Value)5U, Name = "Date Placeholder 4" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties61 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks61 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties61.Append(shapeLocks61);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties76 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape61 = new PlaceholderShape(){ Type = PlaceholderValues.DateAndTime, Size = PlaceholderSizeValues.Half, Index = (UInt32Value)10U };

            applicationNonVisualDrawingProperties76.Append(placeholderShape61);

            nonVisualShapeProperties61.Append(nonVisualDrawingProperties76);
            nonVisualShapeProperties61.Append(nonVisualShapeDrawingProperties61);
            nonVisualShapeProperties61.Append(applicationNonVisualDrawingProperties76);
            ShapeProperties shapeProperties62 = new ShapeProperties();

            TextBody textBody61 = new TextBody();
            A.BodyProperties bodyProperties61 = new A.BodyProperties();
            A.ListStyle listStyle61 = new A.ListStyle();

            A.Paragraph paragraph97 = new A.Paragraph();

            A.Field field23 = new A.Field(){ Id = "{1D065981-0097-4BA9-A692-9FF652AFAF35}", Type = "datetimeFigureOut" };

            A.RunProperties runProperties86 = new A.RunProperties(){ Language = "cs-CZ" };
            runProperties86.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            A.Text text86 = new A.Text();
            text86.Text = "14.03.2023";

            field23.Append(runProperties86);
            field23.Append(text86);
            A.EndParagraphRunProperties endParagraphRunProperties56 = new A.EndParagraphRunProperties(){ Language = "cs-CZ" };

            paragraph97.Append(field23);
            paragraph97.Append(endParagraphRunProperties56);

            textBody61.Append(bodyProperties61);
            textBody61.Append(listStyle61);
            textBody61.Append(paragraph97);

            shape61.Append(nonVisualShapeProperties61);
            shape61.Append(shapeProperties62);
            shape61.Append(textBody61);

            Shape shape62 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties62 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties77 = new NonVisualDrawingProperties(){ Id = (UInt32Value)6U, Name = "Footer Placeholder 5" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties62 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks62 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties62.Append(shapeLocks62);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties77 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape62 = new PlaceholderShape(){ Type = PlaceholderValues.Footer, Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)11U };

            applicationNonVisualDrawingProperties77.Append(placeholderShape62);

            nonVisualShapeProperties62.Append(nonVisualDrawingProperties77);
            nonVisualShapeProperties62.Append(nonVisualShapeDrawingProperties62);
            nonVisualShapeProperties62.Append(applicationNonVisualDrawingProperties77);
            ShapeProperties shapeProperties63 = new ShapeProperties();

            TextBody textBody62 = new TextBody();
            A.BodyProperties bodyProperties62 = new A.BodyProperties();
            A.ListStyle listStyle62 = new A.ListStyle();

            A.Paragraph paragraph98 = new A.Paragraph();
            A.EndParagraphRunProperties endParagraphRunProperties57 = new A.EndParagraphRunProperties(){ Language = "cs-CZ" };

            paragraph98.Append(endParagraphRunProperties57);

            textBody62.Append(bodyProperties62);
            textBody62.Append(listStyle62);
            textBody62.Append(paragraph98);

            shape62.Append(nonVisualShapeProperties62);
            shape62.Append(shapeProperties63);
            shape62.Append(textBody62);

            Shape shape63 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties63 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties78 = new NonVisualDrawingProperties(){ Id = (UInt32Value)7U, Name = "Slide Number Placeholder 6" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties63 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks63 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties63.Append(shapeLocks63);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties78 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape63 = new PlaceholderShape(){ Type = PlaceholderValues.SlideNumber, Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)12U };

            applicationNonVisualDrawingProperties78.Append(placeholderShape63);

            nonVisualShapeProperties63.Append(nonVisualDrawingProperties78);
            nonVisualShapeProperties63.Append(nonVisualShapeDrawingProperties63);
            nonVisualShapeProperties63.Append(applicationNonVisualDrawingProperties78);
            ShapeProperties shapeProperties64 = new ShapeProperties();

            TextBody textBody63 = new TextBody();
            A.BodyProperties bodyProperties63 = new A.BodyProperties();
            A.ListStyle listStyle63 = new A.ListStyle();

            A.Paragraph paragraph99 = new A.Paragraph();

            A.Field field24 = new A.Field(){ Id = "{5D7F2B33-3E35-4F97-A448-9FEB151DF743}", Type = "slidenum" };

            A.RunProperties runProperties87 = new A.RunProperties(){ Language = "cs-CZ" };
            runProperties87.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            A.Text text87 = new A.Text();
            text87.Text = "‹#›";

            field24.Append(runProperties87);
            field24.Append(text87);
            A.EndParagraphRunProperties endParagraphRunProperties58 = new A.EndParagraphRunProperties(){ Language = "cs-CZ" };

            paragraph99.Append(field24);
            paragraph99.Append(endParagraphRunProperties58);

            textBody63.Append(bodyProperties63);
            textBody63.Append(listStyle63);
            textBody63.Append(paragraph99);

            shape63.Append(nonVisualShapeProperties63);
            shape63.Append(shapeProperties64);
            shape63.Append(textBody63);

            shapeTree13.Append(nonVisualGroupShapeProperties13);
            shapeTree13.Append(groupShapeProperties13);
            shapeTree13.Append(shape58);
            shapeTree13.Append(shape59);
            shapeTree13.Append(shape60);
            shapeTree13.Append(shape61);
            shapeTree13.Append(shape62);
            shapeTree13.Append(shape63);

            CommonSlideDataExtensionList commonSlideDataExtensionList13 = new CommonSlideDataExtensionList();

            CommonSlideDataExtension commonSlideDataExtension13 = new CommonSlideDataExtension(){ Uri = "{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}" };

            P14.CreationId creationId13 = new P14.CreationId(){ Val = (UInt32Value)3023498586U };
            creationId13.AddNamespaceDeclaration("p14", "http://schemas.microsoft.com/office/powerpoint/2010/main");

            commonSlideDataExtension13.Append(creationId13);

            commonSlideDataExtensionList13.Append(commonSlideDataExtension13);

            commonSlideData13.Append(shapeTree13);
            commonSlideData13.Append(commonSlideDataExtensionList13);

            ColorMapOverride colorMapOverride12 = new ColorMapOverride();
            A.MasterColorMapping masterColorMapping12 = new A.MasterColorMapping();

            colorMapOverride12.Append(masterColorMapping12);

            slideLayout11.Append(commonSlideData13);
            slideLayout11.Append(colorMapOverride12);

            slideLayoutPart11.SlideLayout = slideLayout11;
        }

        // Generates content of tableStylesPart1.
        private void GenerateTableStylesPart1Content(TableStylesPart tableStylesPart1)
        {
            A.TableStyleList tableStyleList1 = new A.TableStyleList(){ Default = "{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}" };
            tableStyleList1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

            tableStylesPart1.TableStyleList = tableStyleList1;
        }

        // Generates content of viewPropertiesPart1.
        private void GenerateViewPropertiesPart1Content(ViewPropertiesPart viewPropertiesPart1)
        {
            ViewProperties viewProperties1 = new ViewProperties();
            viewProperties1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            viewProperties1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            viewProperties1.AddNamespaceDeclaration("p", "http://schemas.openxmlformats.org/presentationml/2006/main");

            NormalViewProperties normalViewProperties1 = new NormalViewProperties(){ HorizontalBarState = SplitterBarStateValues.Maximized };
            RestoredLeft restoredLeft1 = new RestoredLeft(){ Size = 16993, AutoAdjust = false };
            RestoredTop restoredTop1 = new RestoredTop(){ Size = 94660 };

            normalViewProperties1.Append(restoredLeft1);
            normalViewProperties1.Append(restoredTop1);

            SlideViewProperties slideViewProperties1 = new SlideViewProperties();

            CommonSlideViewProperties commonSlideViewProperties1 = new CommonSlideViewProperties(){ SnapToGrid = false };

            CommonViewProperties commonViewProperties1 = new CommonViewProperties(){ VariableScale = true };

            ScaleFactor scaleFactor1 = new ScaleFactor();
            A.ScaleX scaleX1 = new A.ScaleX(){ Numerator = 118, Denominator = 100 };
            A.ScaleY scaleY1 = new A.ScaleY(){ Numerator = 118, Denominator = 100 };

            scaleFactor1.Append(scaleX1);
            scaleFactor1.Append(scaleY1);
            Origin origin1 = new Origin(){ X = 1014L, Y = 66L };

            commonViewProperties1.Append(scaleFactor1);
            commonViewProperties1.Append(origin1);
            GuideList guideList1 = new GuideList();

            commonSlideViewProperties1.Append(commonViewProperties1);
            commonSlideViewProperties1.Append(guideList1);

            slideViewProperties1.Append(commonSlideViewProperties1);

            NotesTextViewProperties notesTextViewProperties1 = new NotesTextViewProperties();

            CommonViewProperties commonViewProperties2 = new CommonViewProperties();

            ScaleFactor scaleFactor2 = new ScaleFactor();
            A.ScaleX scaleX2 = new A.ScaleX(){ Numerator = 1, Denominator = 1 };
            A.ScaleY scaleY2 = new A.ScaleY(){ Numerator = 1, Denominator = 1 };

            scaleFactor2.Append(scaleX2);
            scaleFactor2.Append(scaleY2);
            Origin origin2 = new Origin(){ X = 0L, Y = 0L };

            commonViewProperties2.Append(scaleFactor2);
            commonViewProperties2.Append(origin2);

            notesTextViewProperties1.Append(commonViewProperties2);
            GridSpacing gridSpacing1 = new GridSpacing(){ Cx = 72008L, Cy = 72008L };

            viewProperties1.Append(normalViewProperties1);
            viewProperties1.Append(slideViewProperties1);
            viewProperties1.Append(notesTextViewProperties1);
            viewProperties1.Append(gridSpacing1);

            viewPropertiesPart1.ViewProperties = viewProperties1;
        }

        // Generates content of extendedFilePropertiesPart1.
        private void GenerateExtendedFilePropertiesPart1Content(ExtendedFilePropertiesPart extendedFilePropertiesPart1)
        {
            Ap.Properties properties1 = new Ap.Properties();
            properties1.AddNamespaceDeclaration("vt", "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes");
            Ap.Template template1 = new Ap.Template();
            template1.Text = "Office Theme 2013 - 2022";
            Ap.TotalTime totalTime1 = new Ap.TotalTime();
            totalTime1.Text = "3";
            Ap.Words words1 = new Ap.Words();
            words1.Text = "0";
            Ap.Application application1 = new Ap.Application();
            application1.Text = "Microsoft Office PowerPoint";
            Ap.PresentationFormat presentationFormat1 = new Ap.PresentationFormat();
            presentationFormat1.Text = "Předvádění na obrazovce (4:3)";
            Ap.Paragraphs paragraphs1 = new Ap.Paragraphs();
            paragraphs1.Text = "0";
            Ap.Slides slides1 = new Ap.Slides();
            slides1.Text = "1";
            Ap.Notes notes1 = new Ap.Notes();
            notes1.Text = "0";
            Ap.HiddenSlides hiddenSlides1 = new Ap.HiddenSlides();
            hiddenSlides1.Text = "0";
            Ap.MultimediaClips multimediaClips1 = new Ap.MultimediaClips();
            multimediaClips1.Text = "0";
            Ap.ScaleCrop scaleCrop1 = new Ap.ScaleCrop();
            scaleCrop1.Text = "false";

            Ap.HeadingPairs headingPairs1 = new Ap.HeadingPairs();

            Vt.VTVector vTVector1 = new Vt.VTVector(){ BaseType = Vt.VectorBaseValues.Variant, Size = (UInt32Value)8U };

            Vt.Variant variant1 = new Vt.Variant();
            Vt.VTLPSTR vTLPSTR1 = new Vt.VTLPSTR();
            vTLPSTR1.Text = "Použitá písma";

            variant1.Append(vTLPSTR1);

            Vt.Variant variant2 = new Vt.Variant();
            Vt.VTInt32 vTInt321 = new Vt.VTInt32();
            vTInt321.Text = "3";

            variant2.Append(vTInt321);

            Vt.Variant variant3 = new Vt.Variant();
            Vt.VTLPSTR vTLPSTR2 = new Vt.VTLPSTR();
            vTLPSTR2.Text = "Motiv";

            variant3.Append(vTLPSTR2);

            Vt.Variant variant4 = new Vt.Variant();
            Vt.VTInt32 vTInt322 = new Vt.VTInt32();
            vTInt322.Text = "1";

            variant4.Append(vTInt322);

            Vt.Variant variant5 = new Vt.Variant();
            Vt.VTLPSTR vTLPSTR3 = new Vt.VTLPSTR();
            vTLPSTR3.Text = "Propojení";

            variant5.Append(vTLPSTR3);

            Vt.Variant variant6 = new Vt.Variant();
            Vt.VTInt32 vTInt323 = new Vt.VTInt32();
            vTInt323.Text = "1";

            variant6.Append(vTInt323);

            Vt.Variant variant7 = new Vt.Variant();
            Vt.VTLPSTR vTLPSTR4 = new Vt.VTLPSTR();
            vTLPSTR4.Text = "Nadpisy snímků";

            variant7.Append(vTLPSTR4);

            Vt.Variant variant8 = new Vt.Variant();
            Vt.VTInt32 vTInt324 = new Vt.VTInt32();
            vTInt324.Text = "1";

            variant8.Append(vTInt324);

            vTVector1.Append(variant1);
            vTVector1.Append(variant2);
            vTVector1.Append(variant3);
            vTVector1.Append(variant4);
            vTVector1.Append(variant5);
            vTVector1.Append(variant6);
            vTVector1.Append(variant7);
            vTVector1.Append(variant8);

            headingPairs1.Append(vTVector1);

            Ap.TitlesOfParts titlesOfParts1 = new Ap.TitlesOfParts();

            Vt.VTVector vTVector2 = new Vt.VTVector(){ BaseType = Vt.VectorBaseValues.Lpstr, Size = (UInt32Value)6U };
            Vt.VTLPSTR vTLPSTR5 = new Vt.VTLPSTR();
            vTLPSTR5.Text = "Arial";
            Vt.VTLPSTR vTLPSTR6 = new Vt.VTLPSTR();
            vTLPSTR6.Text = "Calibri";
            Vt.VTLPSTR vTLPSTR7 = new Vt.VTLPSTR();
            vTLPSTR7.Text = "Calibri Light";
            Vt.VTLPSTR vTLPSTR8 = new Vt.VTLPSTR();
            vTLPSTR8.Text = "Motiv Office";
            Vt.VTLPSTR vTLPSTR9 = new Vt.VTLPSTR();
            vTLPSTR9.Text = "C:\\Users\\matej.kratochvil\\onetable.xlsx!List1!RWATAB";
            Vt.VTLPSTR vTLPSTR10 = new Vt.VTLPSTR();
            vTLPSTR10.Text = "Prezentace aplikace PowerPoint";

            vTVector2.Append(vTLPSTR5);
            vTVector2.Append(vTLPSTR6);
            vTVector2.Append(vTLPSTR7);
            vTVector2.Append(vTLPSTR8);
            vTVector2.Append(vTLPSTR9);
            vTVector2.Append(vTLPSTR10);

            titlesOfParts1.Append(vTVector2);
            Ap.Company company1 = new Ap.Company();
            company1.Text = "";
            Ap.LinksUpToDate linksUpToDate1 = new Ap.LinksUpToDate();
            linksUpToDate1.Text = "false";
            Ap.SharedDocument sharedDocument1 = new Ap.SharedDocument();
            sharedDocument1.Text = "false";
            Ap.HyperlinksChanged hyperlinksChanged1 = new Ap.HyperlinksChanged();
            hyperlinksChanged1.Text = "false";
            Ap.ApplicationVersion applicationVersion1 = new Ap.ApplicationVersion();
            applicationVersion1.Text = "16.0000";

            properties1.Append(template1);
            properties1.Append(totalTime1);
            properties1.Append(words1);
            properties1.Append(application1);
            properties1.Append(presentationFormat1);
            properties1.Append(paragraphs1);
            properties1.Append(slides1);
            properties1.Append(notes1);
            properties1.Append(hiddenSlides1);
            properties1.Append(multimediaClips1);
            properties1.Append(scaleCrop1);
            properties1.Append(headingPairs1);
            properties1.Append(titlesOfParts1);
            properties1.Append(company1);
            properties1.Append(linksUpToDate1);
            properties1.Append(sharedDocument1);
            properties1.Append(hyperlinksChanged1);
            properties1.Append(applicationVersion1);

            extendedFilePropertiesPart1.Properties = properties1;
        }

        private void SetPackageProperties(OpenXmlPackage document)
        {
            document.PackageProperties.Creator = "Matej Kratochvil";
            document.PackageProperties.Title = "Prezentace aplikace PowerPoint";
            document.PackageProperties.Revision = "1";
            document.PackageProperties.Created = System.Xml.XmlConvert.ToDateTime("2023-03-14T00:59:29Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind);
            document.PackageProperties.Modified = System.Xml.XmlConvert.ToDateTime("2023-03-14T01:02:29Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind);
            document.PackageProperties.LastModifiedBy = "Matej Kratochvil";
        }

        #region Binary Data
        private string thumbnailPart1Data = "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";

        private string imagePart1Data = "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AQAAABgAAAAMAAAA8vLyAEwAAABkAAAAAQAAAGsDAAAHBAAAkQMAAAAAAABrAwAACAQAACcAAAAhAPAAAAAAAAAAAAAAAIA/AAAAAAAAAAAAAIA/AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAnAAAAGAAAAAEAAAAAAAAA////AAAAAAAlAAAADAAAAAEAAAAoAAAADAAAAAIAAAAYAAAADAAAAP///wBMAAAAZAAAAAEAAACRAwAABwQAALcDAAAAAAAAkQMAAAgEAAAnAAAAIQDwAAAAAAAAAAAAAACAPwAAAAAAAAAAAACAPwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAJwAAABgAAAACAAAAAAAAAPLy8gAAAAAAJQAAAAwAAAACAAAAKAAAAAwAAAABAAAAGAAAAAwAAADy8vIATAAAAGQAAAABAAAAtwMAAAcEAADdAwAAAAAAALcDAAAIBAAAJwAAACEA8AAAAAAAAAAAAAAAgD8AAAAAAAAAAAAAgD8AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACcAAAAYAAAAAQAAAAAAAAD///8AAAAAACUAAAAMAAAAAQAAACgAAAAMAAAAAgAAABgAAAAMAAAA////AEwAAABkAAAAAQAAAN0DAAAHBAAAAwQAAAAAAADdAwAACAQAACcAAAAhAPAAAAAAAAAAAAAAAIA/AAAAAAAAAAAAAIA/AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAnAAAAGAAAAAIAAAAAAAAA8vLyAAAAAAAlAAAADAAAAAIAAAAoAAAADAAAAAEAAAAYAAAADAAAAPLy8gBMAAAAZAAAAAEAAAADBAAABwQAACkEAAAAAAAAAwQAAAgEAAAnAAAAIQDwAAAAAAAAAAAAAACAPwAAAAAAAAAAAACAPwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAJwAAABgAAAABAAAAAAAAAP///wAAAAAAJQAAAAwAAAABAAAAKAAAAAwAAAACAAAAGAAAAAwAAAD///8ATAAAAGQAAAABAAAAKQQAAAcEAABPBAAAAAAAACkEAAAIBAAAJwAAACEA8AAAAAAAAAAAAAAAgD8AAAAAAAAAAAAAgD8AAAAAAAAAAAAAAAAAAAAA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DgAAAAYAAAAIAAAADgAAAAgAAAAPAAAACAAAAA8AAABUAAAAcAAAAMwBAADDAAAAHgIAAOIAAAACAAAAAAAAAAAAAADMAQAAwwAAAAYAAABMAAAAAAAAAAAAAAAAAAAA//////////9YAAAANAAwACAANAA5ADYADwAAAA8AAAAIAAAADwAAAA8AAAAPAAAAVAAAAFQAAACDAgAAwwAAAJECAADiAAAAAgAAAAAAAAAAAAAAgwIAAMMAAAABAAAATAAAAAAAAAAAAAAAAAAAAP//////////UAAAADAAbAAPAAAAVAAAAFgAAAAYAwAAwwAAAD4DAADiAAAAAgAAAAAAAAAAAAAAGAMAAMMAAAACAAAATAAAAAAAAAAAAAAAAAAAAP//////////UAAAADAAJQAPAAAAGAAAAFQAAACoAAAALwAAAOkAAADdAAAACAEAAAIAAAAAAAAAAAAAAC8AAADpAAAADwAAAEwAAAAAAAAAAAAAAAAAAAD//////////2wAAABjAGEAcwBoACAAYwBvAGwAbABhAHQAZQByAGEAbAAwAA4AAAAPAAAADgAAAA8AAAAIAAAADgAAAA8AAAAGAAAABgAAAA8AAAAIAAAADwAAAAkAAAAPAAAABgAAAFQAAABwAAAAzAEAAOkAAAAeAgAACAEAAAIAAAAAAAAAAAAAAMwBAADpAAAABgAAAEwAAAAAAAAAAAAAAAAAAAD//////////1gAAAAxADUAIAAwADIANQAPAAAADwAAAAgAAAAPAAAADwAAAA8AAABUAAAAVAAAAIMCAADpAAAAkQIAAAgBAAACAAAAAAAAAAAAAACDAgAA6QAAAAEAAABMAAAAAAAAAAAAAAAAAAAA//////////9QAAAAMAAAAA8AAABUAAAAWAAAABgDAADpAAAAPgMAAAgBAAACAAAAAAAAAAAAAAAYAwAA6QAAAAIAAABMAAAAAAAAAAAAAAAAAAAA//////////9QAAAAMAAlAA8AAAAYAAAAJQAAAAwAAAAEAAAAGAAAAAwAAAAAIGAAVAAAALgAAAAaAAAAEAEAACIBAAAvAQAAAgAAAAAAAAAAAAAAGgAAABABAAASAAAATAAAAAAAAAAAAAAAAAAAAP//////////cAAAAEUAeABwAG8AegBpAGMAZQAgAHMAIABSAFcAIAAyADAAIAAlABIAAAAPAAAAEQAAABEAAAAOAAAACAAAAA8AAAAPAAAACAAAAA8AAAAIAAAAFAAAABkAAAAIAAAADwAAAA8AAAAIAAAAGAAAAFQAAABgAAAA8gEAABABAAAeAgAALwEAAAIAAAAAAAAAAAAAAPIBAAAQAQAAAwAAAEwAAAAAAAAAAAAAAAAAAAD//////////1QAAAA3ADQANgD//w8AAAAPAAAADwAAAFQAAABgAAAAZQIAABABAACRAgAALwEAAAIAAAAAAAAAAAAAAGUCAAAQAQAAAwAAAEwAAAAAAAAAAAAAAAAAAAD//////////1QAAAAxADQAOQAAAA8AAAAPAAAADwAAAFQAAABgAAAACQMAABABAAA+AwAALwEAAAIAAAAAAAAAAAAAAAkDAAAQAQAAAwAAAEwAAAAAAAAAAAAAAAAAAAD//////////1QAAAAyADAAJQAAAA8AAAAPAAAAGAAAAFQAAABgAAAAuAMAABABAADtAwAALwEAAAIAAAAAAAAAAAAAALgDAAAQAQAAAwAAAEwAAAAAAAAAAAAAAAAAAAD//////////1QAAAAyADAAJQAAAA8AAAAPAAAAGAAAABgAAAAMAAAAAAAAACUAAAAMAAAABQAAAFQAAACoAAAALwAAADUBAADYAAAAVAEAAAIAAAAAAAAAAAAAAC8AAAA1AQAADwAAAEwAAAAAAAAAAAAAAAAAAAD//////////2wAAAB2AG8BDQFpACAAaQBuAHMAdABpAHQAdQBjAO0AbQAAAA0AAAAPAAAADgAAAAYAAAAIAAAABgAAAA8AAAAOAAAACAAAAAYAAAAIAAAADwAAAA4AAAAGAAAAFgAAAFQAAABUAAAAEAIAADUBAAAeAgAAVAEAAAIAAAAAAAAAAAAAABACAAA1AQAAAQAAAEwAAAAAAAAAAAAAAAAAAAD//////////1AAAAAwAAAADwAAAFQAAABUAAAAgwIAADUBAACRAgAAVAEAAAIAAAAAAAAAAAAAAIMCAAA1AQAAAQAAAEwAAAAAAAAAAAAAAAAAAAD//////////1AAAAAwAAAADwAAAFQAAACoAAAALwAAAFsBAADlAAAAegEAAAIAAAAAAAAAAAAAAC8AAABbAQAADwAAAEwAAAAAAAAAAAAAAAAAAAD//////////2wAAABrAHIA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AgAAAAAAAAAAAAAA8gEAAEACAAADAAAATAAAAAAAAAAAAAAAAAAAAP//////////VAAAADEAMAA5ADAADwAAAA8AAAAPAAAAVAAAAFgAAAB0AgAAQAIAAJECAABfAgAAAgAAAAAAAAAAAAAAdAIAAEACAAACAAAATAAAAAAAAAAAAAAAAAAAAP//////////UAAAADUAMwAPAAAADwAAAFQAAABgAAAACQMAAEACAAA+AwAAXwIAAAIAAAAAAAAAAAAAAAkDAABAAgAAAwAAAEwAAAAAAAAAAAAAAAAAAAD//////////1QAAAAzADgAJQAAAA8AAAAPAAAAGAAAAFQAAABgAAAAuAMAAEACAADtAwAAXwIAAAIAAAAAAAAAAAAAALgDAABAAgAAAwAAAEwAAAAAAAAAAAAAAAAAAAD//////////1QAAAA0ADkAJQAAAA8AAAAPAAAAGAAAABgAAAAMAAAAAAAAACUAAAAMAAAABQAAAFQAAACoAAAALwAAAGUCAADYAAAAhAIAAAIAAAAAAAAAAAAAAC8AAABlAgAADwAAAEwAAAAAAAAAAAAAAAAAAAD//////////2wAAAB2AG8BDQFpACAAaQBuAHMAdABpAHQAdQBjAO0AbQAAAA0AAAAPAAAADgAAAAYAAAAIAAAABgAAAA8AAAAOAAAACAAAAAYAAAAIAAAADwAAAA4AAAAGAAAAFgAAAFQAAACoAAAALwAAAIsCAADlAAAAqgIAAAIAAAAAAAAAAAAAAC8AAACLAgAADwAAAEwAAAAAAAAAAAAAAAAAAAD//////////2wAAABrAHIAeQB0AOkAIABkAGwAdQBoAG8AcABpAHMAeQAAAA4AAAAJAAAADgAAAAgAAAAPAAAACAAAAA8AAAAGAAAADwAAAA8AAAAPAAAADwAAAAYAAAAOAAAADgAAAFQAAADkAAAALwAAALECAABfAQAA0AIAAAIAAAAAAAAAAAAAAC8AAACxAgAAGQAAAEwAAAAAAAAAAAAAAAAAAAD//////////4AAAAB6AGEAagAuACAAbwBiAGMAaABvAGQAbgDtACAAbgBlAG0AbwB2AGkAdABvAHMAdADtAAAADQAAAA8AAAAGAAAACAAAAAgAAAAPAAAADwAAAA4AAAAPAAAADwAAAA8AAAAPAAAABgAAAAgAAAAPAAAA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CAAAABQAAAAZAAAACAAAAA8AAAAPAAAADwAAAAgAAAAYAAAACAAAABQAAAASAAAAEgAAABMAAAAIAAAAEQAAAFQAAABsAAAA2wEAAHoEAAAeAgAAmQQAAAIAAAAAAAAAAAAAANsBAAB6BAAABQAAAEwAAAAAAAAAAAAAAAAAAAD//////////1gAAAAyACAANQAwADYAAAAPAAAACAAAAA8AAAAPAAAADwAAAFQAAABsAAAATgIAAHoEAACRAgAAmQQAAAIAAAAAAAAAAAAAAE4CAAB6BAAABQAAAEwAAAAAAAAAAAAAAAAAAAD//////////1gAAAAxACAANQA1ADUAIAAPAAAACAAAAA8AAAAPAAAADwAAAFQAAABkAAAA+gIAAHoEAAA+AwAAmQQAAAIAAAAAAAAAAAAAAPoCAAB6BAAABAAAAEwAAAAAAAAAAAAAAAAAAAD//////////1QAAAAxADAAMAAlAA8AAAAPAAAADwAAABgAAABUAAAAYAAAALgDAAB6BAAA7QMAAJkEAAACAAAAAAAAAAAAAAC4AwAAegQAAAMAAABMAAAAAAAAAAAAAAAAAAAA//////////9UAAAAMwA3ACUAAAAPAAAADwAAABgAAAAYAAAADAAAAAAAAAAlAAAADAAAAAUAAABUAAAACAEAAC8AAACfBAAAuQEAAL4EAAACAAAAAAAAAAAAAAAvAAAAnwQAAB8AAABMAAAAAAAAAAAAAAAAAAAA//////////+MAAAAegBhAGoALgAgAG8AYgB5AHQAbgBvAHUAIABuAGUAbQBvAHYAaQB0AG8AcwB0AO0AIABSAEUAVABBAEkATAAAAA0AAAAPAAAABgAAAAgAAAAIAAAADwAAAA8AAAAOAAAACAAAAA8AAAAPAAAADwAAAAgAAAAPAAAADwAAABYAAAAPAAAADQAAAAYAAAAIAAAADwAAAA4AAAAIAAAABgAAAAgAAAAUAAAAEgAAABAAAAASAAAACAAAAA8AAABUAAAArAAAAC8AAADFBAAAAAEAAOQEAAACAAAAAAAAAAAAAAAvAAAAxQQAABAAAABMAAAAAAAAAAAAAAAAAAAA//////////9sAAAAdgAgAHMAZQBsAGgA4QBuAO0AIABSAEUAVABBAEkATAANAAAACAAAAA4AAAAPAAAABgAAAA8AAAAPAAAADwAAAAYAAAAIAAAAFAAAABIAAAAQAAAAEgAAAAgAAAAPAAAAJQAAAAwAAAAEAAAAGAAAAAwAAAAAIGAAVAAAAMAAAAAaAAAA7AQAADEBAAALBQAAAgAAAAAAAAAAAAAAGgAAAOwEAAATAAAATAAAAAAAAAAAAAAAAAAAAP//////////dAAAAEUAeABwAG8AegBpAGMAZQAgAHMAIABSAFcAIAAxADUAMAAgACUAAAASAAAADwAAABEAAAARAAAADgAAAAgAAAAPAAAADwAAAAgAAAAPAAAACAAAABQAAAAZAAAACAAAAA8AAAAPAAAADwAAAAgAAAAYAAAAVAAAAGwAAADbAQAA7AQAAB4CAAALBQAAAgAAAAAAAAAAAAAA2wEAAOwEAAAFAAAATAAAAAAAAAAAAAAAAAAAAP//////////WAAAADUAIAA0ADcANwAAAA8AAAAIAAAADwAAAA8AAAAPAAAAVAAAAGwAAABOAgAA7AQAAJECAAALBQAAAgAAAAAAAAAAAAAATgIAAOwEAAAFAAAATAAAAAAAAAAAAAAAAAAAAP//////////WAAAADcAIAA0ADAAMQAAAA8AAAAIAAAADwAAAA8AAAAPAAAAVAAAAGQAAAD6AgAA7AQAAD4DAAALBQAAAgAAAAAAAAAAAAAA+gIAAOwEAAAEAAAATAAAAAAAAAAAAAAAAAAAAP//////////VAAAADEANAA2ACUADwAAAA8AAAAPAAAAGAAAAFQAAABgAAAAuAMAAOwEAADtAwAACwUAAAIAAAAAAAAAAAAAALgDAADsBAAAAwAAAEwAAAAAAAAAAAAAAAAAAAD//////////1QAAAA3ADAAJQAAAA8AAAAPAAAAGAAAABgAAAAMAAAAAAAAACUAAAAMAAAABQAAAFQAAACoAAAALwAAABEFAADeAAAAMAUAAAIAAAAAAAAAAAAAAC8AAAARBQAADwAAAEwAAAAAAAAAAAAAAAAAAAD//////////2wAAABoAGkAZwBoACAAcgBpAHMAawAgAPoAdgAbAXIAeQAAAA8AAAAGAAAADwAAAA8AAAAIAAAACQAAAAYAAAAOAAAADgAAAAgAAAAPAAAADQAAAA8AAAAJAAAADgAAAFQAAADAAAAALwAAADcFAAAPAQAAVgUAAAIAAAAAAAAAAAAAAC8AAAA3BQAAEwAAAEwAAAAAAAAAAAAAAAAAAAD//////////3QAAABoAGkAZwBoACAAcgBpAHMAawAgAGQAbAB1AGgAbwBwAGkAcwB5AAAADwAAAAYAAAAPAAAADwAAAAgAAAAJAAAABgAAAA4AAAAOAAAACAAAAA8AAAAGAAAADwAAAA8AAAAPAAAADwAAAAYAAAAOAAAADgAAAFQAAACEAAAALwAAAF0FAACZAAAAfAUAAAIAAAAAAAAAAAAAAC8AAABdBQAACQAAAEwAAAAAAAAAAAAAAAAAAAD//////////2AAAAB2ACAAcwBlAGwAaADhAG4A7QAAAA0AAAAIAAAADgAAAA8AAAAGAAAADwAAAA8AAAAPAAAABgAAACUAAAAMAAAABAAAABgAAAAMAAAAACBgAFQAAADoAAAAGgAAAIQFAACdAQAAowUAAAIAAAAAAAAAAAAAABoAAACEBQAAGgAAAEwAAAAAAAAAAAAAAAAAAAD//////////4AAAABFAHgAcABvAHoAaQBjAGUAIABzACAAUgBXACAAMQA1ADAAIAAlACAAUgBFAFQAQQBJAEwAEgAAAA8AAAARAAAAEQAAAA4AAAAIAAAADwAAAA8AAAAIAAAADwAAAAgAAAAUAAAAGQAAAAgAAAAPAAAADwAAAA8AAAAIAAAAGAAAAAgAAAAUAAAAEgAAABIAAAATAAAACAAAABEAAABUAAAAVAAAABACAACEBQAAHgIAAKMFAAACAAAAAAAAAAAAAAAQAgAAhAUAAAEAAABMAAAAAAAAAAAAAAAAAAAA//////////9QAAAAMABgAA8AAABUAAAAVAAAAIMCAACEBQAAkQIAAKMFAAACAAAAAAAAAAAAAACDAgAAhAUAAAEAAABMAAAAAAAAAAAAAAAAAAAA//////////9QAAAAMABvAA8AAAAYAAAADAAAAAAAAAAlAAAADAAAAAUAAABUAAAArAAAAC8AAACpBQAAAAEAAMgFAAACAAAAAAAAAAAAAAAvAAAAqQUAABAAAABMAAAAAAAAAAAAAAAAAAAA//////////9sAAAAdgAgAHMAZQBsAGgA4QBuAO0AIABSAEUAVABBAEkATAANAAAACAAAAA4AAAAPAAAABgAAAA8AAAAPAAAADwAAAAYAAAAIAAAAFAAAABIAAAAQAAAAEgAAAAgAAAAPAAAA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AQAAAP////8hAPAAAAAAAAAAAAAAAIA/AAAAAAAAAAAAAIA/AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAlAAAADAAAAAEAAAAbAAAAEAAAADgCAAAAAAAANgAAABAAAAA4AgAA/////yUAAAAMAAAACQAAAEwAAABkAAAAAAAAAAAAAAD//////////zgCAAAAAAAAAQAAAP////8hAPAAAAAAAAAAAAAAAIA/AAAAAAAAAAAAAIA/AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAlAAAADAAAAAEAAAAbAAAAEAAAAKsCAAAAAAAANgAAABAAAACrAgAA/////yUAAAAMAAAACQAAAEwAAABkAAAAAAAAAAAAAAD//////////6sCAAAAAAAAAQAAAP////8hAPAAAAAAAAAAAAAAAIA/AAAAAAAAAAAAAIA/AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAlAAAADAAAAAEAAAAbAAAAEAAAAFgDAAAAAAAANgAAABAAAABYAwAA/////yUAAAAMAAAACQAAAEwAAABkAAAAAAAAAAAAAAD//////////1gDAAAAAAAAAQAAAP////8hAPAAAAAAAAAAAAAAAIA/AAAAAAAAAAAAAIA/AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAnAAAAGAAAAAoAAAAAAAAAAAAAAAAAAAAlAAAADAAAAAoAAAAYAAAADAAAAAAAAAAZAAAADAAAAP///wBMAAAAZAAAAAAAAAAAAAAABwQAAAAAAAAAAAAA/////wgEAAACAAAAIQDwAAAAAAAAAAAAAACAPwAAAAAAAAAAAACAPwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAJQAAAAwAAAAIAAAAKAAAAAwAAAAKAAAAGAAAAAwAAADg4OAAGQAAAAwAAADg4OAAJQAAAAwAAAABAAAAGwAAABAAAAAHBAAAAAAAADYAAAAQAAAABwQAAP////8lAAAADAAAAAkAAABMAAAAZAAAAAAAAAAAAAAA//////////8HBAAAAAAAAAEAAAD/////IQDwAAAAAAAAAAAAAACAPwAAAAAAAAAAAACAPwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAJwAAABgAAAAKAAAAAAAAAKampgAAAAAAJQAAAAwAAAAKAAAAGAAAAAwAAACmpqYAGQAAAAwAAAD///8AKAAAAAwAAAABAAAAJgAAABwAAAABAAAAAAAAAAAAAAAAAAAApqamACUAAAAMAAAAAQAAABsAAAAQAAAAxQEAAAEAAAA2AAAAEAAAAMUBAAAmAAAAJQAAAAwAAAAJAAAATAAAAGQAAADFAQAAAQAAAMUBAAAlAAAAxQEAAAEAAAABAAAAJQAAACEA8AAAAAAAAAAAAAAAgD8AAAAAAAAAAAAAgD8AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACUAAAAMAAAAAQAAABsAAAAQAAAAqwIAAAEAAAA2AAAAEAAAAKsCAAAmAAAAJQAAAAwAAAAJAAAATAAAAGQAAACrAgAAAQAAAKsCAAAlAAAAqwIAAAEAAAABAAAAJQAAACEA8AAAAAAAAAAAAAAAgD8AAAAAAAAAAAAAgD8AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACcAAAAYAAAACwAAAAAAAAAAAAAAAAAAACUAAAAMAAAACwAAACgAAAAMAAAACgAAABgAAAAMAAAAAAAAAEwAAABkAAAAAAAAACYAAAAHBAAAJwAAAAAAAAAmAAAACAQAAAIAAAAhAPAAAAAAAAAAAAAAAIA/AAAAAAAAAAAAAIA/AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAnAAAAGAAAAAoAAAAAAAAApqamAAAAAAAlAAAADAAAAAoAAAAoAAAADAAAAAsAAAAYAAAADAAAAKampgAlAAAADAAAAAEAAAAbAAAAEAAAAMUBAAAoAAAANgAAABAAAADFAQAACwEAACUAAAAMAAAACQAAAEwAAABkAAAAxQEAACgAAADFAQAACgEAAMUBAAAoAAAAAQAAAOMAAAAhAPAAAAAAAAAAAAAAAIA/AAAAAAAAAAAAAIA/AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAlAAAADAAAAAEAAAAbAAAAEAAAAKsCAAAoAAAANgAAABAAAACrAgAACwEAACUAAAAMAAAACQAAAEwAAABkAAAAqwIAACgAAACrAgAACgEAAKsCAAAoAAAAAQAAAOMAAAAhAPAAAAAAAAAAAAAAAIA/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AAAAAAAAAAAAAIA/AAAAAAAAAAAAAIA/AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAnAAAAGAAAAAsAAAAAAAAAAAAAAAAAAAAlAAAADAAAAAsAAAAYAAAADAAAAAAAAAAZAAAADAAAAP///wBMAAAAZAAAAAAAAAA+BgAABwQAAD8GAAAAAAAAPgYAAAgEAAACAAAAIQDwAAAAAAAAAAAAAACAPwAAAAAAAAAAAACAPwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAJQAAAAwAAAAIAAAAKAAAAAwAAAALAAAAGAAAAAwAAADg4OAAGQAAAAwAAADg4OAAJQAAAAwAAAABAAAAGwAAABAAAAAHBAAAGQYAADYAAAAQAAAABwQAAD4GAAAlAAAADAAAAAkAAABMAAAAZAAAAAcEAAAZBgAABwQAAD0GAAAHBAAAGQYAAAEAAAAlAAAAIQDwAAAAAAAAAAAAAACAPwAAAAAAAAAAAACAPwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAJQAAAAwAAAABAAAAGwAAABAAAAAAAAAAQAYAADYAAAAQAAAAAAAAAEEGAAAlAAAADAAAAAkAAABMAAAAZAAAAAAAAAAAAAAA//////////8AAAAAQAYAAAEAAAABAAAAIQDwAAAAAAAAAAAAAACAPwAAAAAAAAAAAACAPwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAJQAAAAwAAAABAAAAGwAAABAAAADFAQAAQAYAADYAAAAQAAAAxQEAAEEGAAAlAAAADAAAAAkAAABMAAAAZAAAAAAAAAAAAAAA///////////FAQAAQAYAAAEAAAABAAAAIQDwAAAAAAAAAAAAAACAPwAAAAAAAAAAAACAPwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAJQAAAAwAAAABAAAAGwAAABAAAAA4AgAAQAYAADYAAAAQAAAAOAIAAEEGAAAlAAAADAAAAAkAAABMAAAAZAAAAAAAAAAAAAAA//////////84AgAAQAYAAAEAAAABAAAAIQDwAAAAAAAAAAAAAACAPwAAAAAAAAAAAACAPwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAJQAAAAwAAAABAAAAGwAAABAAAACrAgAAQAYAADYAAAAQAAAAqwIAAEEGAAAlAAAADAAAAAkAAABMAAAAZAAAAAAAAAAAAAAA//////////+rAgAAQAYAAAEAAAABAAAAIQDwAAAAAAAAAAAAAACAPwAAAAAAAAAAAACAPwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAJQAAAAwAAAABAAAAGwAAABAAAABYAwAAQAYAADYAAAAQAAAAWAMAAEEGAAAlAAAADAAAAAkAAABMAAAAZAAAAAAAAAAAAAAA//////////9YAwAAQAYAAAEAAAABAAAAIQDwAAAAAAAAAAAAAACAPwAAAAAAAAAAAACAPwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAJQAAAAwAAAABAAAAGwAAABAAAAAHBAAAQAYAADYAAAAQAAAABwQAAEEGAAAlAAAADAAAAAkAAABMAAAAZAAAAAAAAAAAAAAA//////////8HBAAAQAYAAAEAAAABAAAAIQDwAAAAAAAAAAAAAACAPwAAAAAAAAAAAACAPwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAJQAAAAwAAAABAAAAGwAAABAAAAAIBAAAAAAAADYAAAAQAAAACQQAAAAAAAAlAAAADAAAAAkAAABMAAAAZAAAAAAAAAAAAAAA//////////8IBAAAAAAAAAEAAAABAAAAIQDwAAAAAAAAAAAAAACAPwAAAAAAAAAAAACAPwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAJQAAAAwAAAABAAAAGwAAABAAAAAIBAAAJwAAADYAAAAQAAAACQQAACcAAAAlAAAADAAAAAkAAABMAAAAZAAAAAAAAAAAAAAA//////////8IBAAAJwAAAAEAAAABAAAAIQDwAAAAAAAAAAAAAACAPwAAAAAAAAAAAACAPwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAJQAAAAwAAAABAAAAGwAAABAAAAAIBAAATQAAADYAAAAQAAAACQQAAE0AAAAlAAAADAAAAAkAAABMAAAAZAAAAAAAAAAAAAAA//////////8IBAAATQAAAAEAAAABAAAAIQDwAAAAAAAAAAAAAACAPwAAAAAAAAAAAACAPwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAJQAAAAwAAAABAAAAGwAAABAAAAAIBAAAcwAAADYAAAAQAAAACQQAAHMAAAAlAAAADAAAAAkAAABMAAAAZAAAAAAAAAAAAAAA//////////8IBAAAcwAAAAEAAAABAAAAIQDwAAAAAAAAAAAAAACAPwAAAAAAAAAAAACAPwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAJQAAAAwAAAABAAAAGwAAABAAAAAIBAAAmQAAADYAAAAQAAAACQQAAJkAAAAlAAAADAAAAAkAAABMAAAAZAAAAAAAAAAAAAAA//////////8IBAAAmQAAAAEAAAABAAAAIQDwAAAAAAAAAAAAAACAPwAAAAAAAAAAAACAPwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAJQAAAAwAAAABAAAAGwAAABAAAAAIBAAAvwAAADYAAAAQAAAACQQAAL8AAAAlAAAADAAAAAkAAABMAAAAZAAAAAAAAAAAAAAA//////////8IBAAAvwAAAAEAAAABAAAAIQDwAAAAAAAAAAAAAACAPwAAAAAAAAAAAACAPwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAJQAAAAwAAAABAAAAGwAAABAAAAAIBAAA5QAAADYAAAAQAAAACQQAAOUAAAAlAAAADAAAAAkAAABMAAAAZAAAAAAAAAAAAAAA//////////8IBAAA5QAAAAEAAAABAAAAIQDwAAAAAAAAAAAAAACAPwAAAAAAAAAAAACAPwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAJQAAAAwAAAABAAAAGwAAABAAAAAIBAAACwEAADYAAAAQAAAACQQAAAsBAAAlAAAADAAAAAkAAABMAAAAZAAAAAAAAAAAAAAA//////////8IBAAACwEAAAEAAAABAAAAIQDwAAAAAAAAAAAAAACAPwAAAAAAAAAAAACAPwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAJQAAAAwAAAABAAAAGwAAABAAAAAIBAAAMQEAADYAAAAQAAAACQQAADEBAAAlAAAADAAAAAkAAABMAAAAZAAAAAAAAAAAAAAA//////////8IBAAAMQEAAAEAAAABAAAAIQDwAAAAAAAAAAAAAACAPwAAAAAAAAAAAACAPwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAJQAAAAwAAAABAAAAGwAAABAAAAAIBAAAVwEAADYAAAAQAAAACQQAAFcBAAAlAAAADAAAAAkAAABMAAAAZAAAAAAAAAAAAAAA//////////8IBAAAVwEAAAEAAAABAAAAIQDwAAAAAAAAAAAAAACAPwAAAAAAAAAAAACAPwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAJQAAAAwAAAABAAAAGwAAABAAAAAIBAAAfQEAADYAAAAQAAAACQQAAH0BAAAlAAAADAAAAAkAAABMAAAAZAAAAAAAAAAAAAAA//////////8IBAAAfQEAAAEAAAABAAAAIQDwAAAAAAAAAAAAAACAPwAAAAAAAAAAAACAPwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAJQAAAAwAAAABAAAAGwAAABAAAAAIBAAAowEAADYAAAAQAAAACQQAAKMBAAAlAAAADAAAAAkAAABMAAAAZAAAAAAAAAAAAAAA//////////8IBAAAowEAAAEAAAABAAAAIQDwAAAAAAAAAAAAAACAPwAAAAAAAAAAAACAPwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAJQAAAAwAAAABAAAAGwAAABAAAAAIBAAAyQEAADYAAAAQAAAACQQAAMkBAAAlAAAADAAAAAkAAABMAAAAZAAAAAAAAAAAAAAA//////////8IBAAAyQEAAAEAAAABAAAAIQDwAAAAAAAAAAAAAACAPwAAAAAAAAAAAACAPwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAJQAAAAwAAAABAAAAGwAAABAAAAAIBAAA7wEAADYAAAAQAAAACQQAAO8BAAAlAAAADAAAAAkAAABMAAAAZAAAAAAAAAAAAAAA//////////8IBAAA7wEAAAEAAAABAAAAIQDwAAAAAAAAAAAAAACAPwAAAAAAAAAAAACAPwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAJQAAAAwAAAABAAAAGwAAABAAAAAIBAAAFQIAADYAAAAQAAAACQQAABUCAAAlAAAADAAAAAkAAABMAAAAZAAAAAAAAAAAAAAA//////////8IBAAAFQIAAAEAAAABAAAAIQDwAAAAAAAAAAAAAACAPwAAAAAAAAAAAACAPwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAJQAAAAwAAAABAAAAGwAAABAAAAAIBAAAOwIAADYAAAAQAAAACQQAADsCAAAlAAAADAAAAAkAAABMAAAAZAAAAAAAAAAAAAAA//////////8IBAAAOwIAAAEAAAABAAAAIQDwAAAAAAAAAAAAAACAPwAAAAAAAAAAAACAPwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAJQAAAAwAAAABAAAAGwAAABAAAAAIBAAAYQIAADYAAAAQAAAACQQAAGECAAAlAAAADAAAAAkAAABMAAAAZAAAAAAAAAAAAAAA//////////8IBAAAYQIAAAEAAAABAAAAIQDwAAAAAAAAAAAAAACAPwAAAAAAAAAAAACAPwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAJQAAAAwAAAABAAAAGwAAABAAAAAIBAAAhwIAADYAAAAQAAAACQQAAIcCAAAlAAAADAAAAAkAAABMAAAAZAAAAAAAAAAAAAAA//////////8IBAAAhwIAAAEAAAABAAAAIQDwAAAAAAAAAAAAAACAPwAAAAAAAAAAAACAPwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAJQAAAAwAAAABAAAAGwAAABAAAAAIBAAArQIAADYAAAAQAAAACQQAAK0CAAAlAAAADAAAAAkAAABMAAAAZAAAAAAAAAAAAAAA//////////8IBAAArQIAAAEAAAABAAAAIQDwAAAAAAAAAAAAAACAPwAAAAAAAAAAAACAPwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAJQAAAAwAAAABAAAAGwAAABAAAAAIBAAA0wIAADYAAAAQAAAACQQAANMCAAAlAAAADAAAAAkAAABMAAAAZAAAAAAAAAAAAAAA//////////8IBAAA0wIAAAEAAAABAAAAIQDwAAAAAAAAAAAAAACAPwAAAAAAAAAAAACAPwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAJQAAAAwAAAABAAAAGwAAABAAAAAIBAAA+QIAADYAAAAQAAAACQQAAPkCAAAlAAAADAAAAAkAAABMAAAAZAAAAAAAAAAAAAAA//////////8IBAAA+QIAAAEAAAABAAAAIQDwAAAAAAAAAAAAAACAPwAAAAAAAAAAAACAPwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAJQAAAAwAAAABAAAAGwAAABAAAAAIBAAAHwMAADYAAAAQAAAACQQAAB8DAAAlAAAADAAAAAkAAABMAAAAZAAAAAAAAAAAAAAA//////////8IBAAAHwMAAAEAAAABAAAAIQDwAAAAAAAAAAAAAACAPwAAAAAAAAAAAACAPwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAJQAAAAwAAAABAAAAGwAAABAAAAAIBAAARQMAADYAAAAQAAAACQQAAEUDAAAlAAAADAAAAAkAAABMAAAAZAAAAAAAAAAAAAAA//////////8IBAAARQMAAAEAAAABAAAAIQDwAAAAAAAAAAAAAACAPwAAAAAAAAAAAACAPwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAJQAAAAwAAAABAAAAGwAAABAAAAAIBAAAawMAADYAAAAQAAAACQQAAGsDAAAlAAAADAAAAAkAAABMAAAAZAAAAAAAAAAAAAAA//////////8IBAAAawMAAAEAAAABAAAAIQDwAAAAAAAAAAAAAACAPwAAAAAAAAAAAACAPwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAJQAAAAwAAAABAAAAGwAAABAAAAAIBAAAkQMAADYAAAAQAAAACQQAAJEDAAAlAAAADAAAAAkAAABMAAAAZAAAAAAAAAAAAAAA//////////8IBAAAkQMAAAEAAAABAAAAIQDwAAAAAAAAAAAAAACAPwAAAAAAAAAAAACAPwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAJQAAAAwAAAABAAAAGwAAABAAAAAIBAAAtwMAADYAAAAQAAAACQQAALcDAAAlAAAADAAAAAkAAABMAAAAZAAAAAAAAAAAAAAA//////////8IBAAAtwMAAAEAAAABAAAAIQDwAAAAAAAAAAAAAACAPwAAAAAAAAAAAACAPwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAJQAAAAwAAAABAAAAGwAAABAAAAAIBAAA3QMAADYAAAAQAAAACQQAAN0DAAAlAAAADAAAAAkAAABMAAAAZAAAAAAAAAAAAAAA//////////8IBAAA3QMAAAEAAAABAAAAIQDwAAAAAAAAAAAAAACAPwAAAAAAAAAAAACAPwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAJQAAAAwAAAABAAAAGwAAABAAAAAIBAAAAwQAADYAAAAQAAAACQQAAAMEAAAlAAAADAAAAAkAAABMAAAAZAAAAAAAAAAAAAAA//////////8IBAAAAwQAAAEAAAABAAAAIQDwAAAAAAAAAAAAAACAPwAAAAAAAAAAAACAPwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAJQAAAAwAAAABAAAAGwAAABAAAAAIBAAAKQQAADYAAAAQAAAACQQAACkEAAAlAAAADAAAAAkAAABMAAAAZAAAAAAAAAAAAAAA//////////8IBAAAKQQAAAEAAAABAAAAIQDwAAAAAAAAAAAAAACAPwAAAAAAAAAAAACAPwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAJQAAAAwAAAABAAAAGwAAABAAAAAIBAAATwQAADYAAAAQAAAACQQAAE8EAAAlAAAADAAAAAkAAABMAAAAZAAAAAAAAAAAAAAA//////////8IBAAATwQAAAEAAAABAAAAIQDwAAAAAAAAAAAAAACAPwAAAAAAAAAAAACAPwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAJQAAAAwAAAABAAAAGwAAABAAAAAIBAAAdQQAADYAAAAQAAAACQQAAHUEAAAlAAAADAAAAAkAAABMAAAAZAAAAAAAAAAAAAAA//////////8IBAAAdQQAAAEAAAABAAAAIQDwAAAAAAAAAAAAAACAPwAAAAAAAAAAAACAPwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAJQAAAAwAAAABAAAAGwAAABAAAAAIBAAAmwQAADYAAAAQAAAACQQAAJsEAAAlAAAADAAAAAkAAABMAAAAZAAAAAAAAAAAAAAA//////////8IBAAAmwQAAAEAAAABAAAAIQDwAAAAAAAAAAAAAACAPwAAAAAAAAAAAACAPwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAJQAAAAwAAAABAAAAGwAAABAAAAAIBAAAwQQAADYAAAAQAAAACQQAAMEEAAAlAAAADAAAAAkAAABMAAAAZAAAAAAAAAAAAAAA//////////8IBAAAwQQAAAEAAAABAAAAIQDwAAAAAAAAAAAAAACAPwAAAAAAAAAAAACAPwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAJQAAAAwAAAABAAAAGwAAABAAAAAIBAAA5wQAADYAAAAQAAAACQQAAOcEAAAlAAAADAAAAAkAAABMAAAAZAAAAAAAAAAAAAAA//////////8IBAAA5wQAAAEAAAABAAAAIQDwAAAAAAAAAAAAAACAPwAAAAAAAAAAAACAPwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAJQAAAAwAAAABAAAAGwAAABAAAAAIBAAADQUAADYAAAAQAAAACQQAAA0FAAAlAAAADAAAAAkAAABMAAAAZAAAAAAAAAAAAAAA//////////8IBAAADQUAAAEAAAABAAAAIQDwAAAAAAAAAAAAAACAPwAAAAAAAAAAAACAPwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAJQAAAAwAAAABAAAAGwAAABAAAAAIBAAAMwUAADYAAAAQAAAACQQAADMFAAAlAAAADAAAAAkAAABMAAAAZAAAAAAAAAAAAAAA//////////8IBAAAMwUAAAEAAAABAAAAIQDwAAAAAAAAAAAAAACAPwAAAAAAAAAAAACAPwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAJQAAAAwAAAABAAAAGwAAABAAAAAIBAAAWQUAADYAAAAQAAAACQQAAFkFAAAlAAAADAAAAAkAAABMAAAAZAAAAAAAAAAAAAAA//////////8IBAAAWQUAAAEAAAABAAAAIQDwAAAAAAAAAAAAAACAPwAAAAAAAAAAAACAPwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAJQAAAAwAAAABAAAAGwAAABAAAAAIBAAAfwUAADYAAAAQAAAACQQAAH8FAAAlAAAADAAAAAkAAABMAAAAZAAAAAAAAAAAAAAA//////////8IBAAAfwUAAAEAAAABAAAAIQDwAAAAAAAAAAAAAACAPwAAAAAAAAAAAACAPwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAJQAAAAwAAAABAAAAGwAAABAAAAAIBAAApQUAADYAAAAQAAAACQQAAKUFAAAlAAAADAAAAAkAAABMAAAAZAAAAAAAAAAAAAAA//////////8IBAAApQUAAAEAAAABAAAAIQDwAAAAAAAAAAAAAACAPwAAAAAAAAAAAACAPwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAJQAAAAwAAAABAAAAGwAAABAAAAAIBAAAywUAADYAAAAQAAAACQQAAMsFAAAlAAAADAAAAAkAAABMAAAAZAAAAAAAAAAAAAAA//////////8IBAAAywUAAAEAAAABAAAAIQDwAAAAAAAAAAAAAACAPwAAAAAAAAAAAACAPwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAJQAAAAwAAAABAAAAGwAAABAAAAAIBAAA8QUAADYAAAAQAAAACQQAAPEFAAAlAAAADAAAAAkAAABMAAAAZAAAAAAAAAAAAAAA//////////8IBAAA8QUAAAEAAAABAAAAIQDwAAAAAAAAAAAAAACAPwAAAAAAAAAAAACAPwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAJQAAAAwAAAABAAAAGwAAABAAAAAIBAAAGAYAADYAAAAQAAAACQQAABgGAAAlAAAADAAAAAkAAABMAAAAZAAAAAAAAAAAAAAA//////////8IBAAAGAYAAAEAAAABAAAAIQDwAAAAAAAAAAAAAACAPwAAAAAAAAAAAACAPwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAJQAAAAwAAAABAAAAGwAAABAAAAAIBAAAPwYAADYAAAAQAAAACQQAAD8GAAAlAAAADAAAAAkAAABMAAAAZAAAAAAAAAAAAAAA//////////8IBAAAPwYAAAEAAAABAAAAIQDwAAAAAAAAAAAAAACAPwAAAAAAAAAAAACAPwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAJwAAABgAAAALAAAAAAAAAAAAAAAAAAAAJQAAAAwAAAALAAAAJQAAAAwAAAANAACAIgAAAAwAAAD/////IQAAAAgAAAAlAAAADAAAAAYAAAAZAAAADAAAAODg4AAYAAAADAAAAODg4AAeAAAAGAAAAAEAAAABAAAACAQAAEAGAAAeAAAAGAAAAAEAAAABAAAACAQAAEAGAABGAAAAQAAAADQAAABFTUYrKkAAACQAAAAYAAAAAACAPwAAAIAAAACAAACAPwAAAIAAAACABEAAAAwAAAAAAAAARgAAAKQAAACYAAAARU1GKypAAAAkAAAAGAAAAAAAgD8AAAAAAAAAAAAAgD8AAAAAAAAAADJAAAEcAAAAEAAAAAAAgD8AAIA/AOCARADgx0QqQAAAJAAAABgAAAAAAIA/AAAAgAAAAIAAAIA/AAAAgAAAAIAIQAAEGAAAAAwAAAACEMDbAAAAAAMAABA0QAAADAAAAAAAAAAEQAAADAAAAAAAAABLAAAAEAAAAAAAAAAFAAAAIgAAAAwAAAD/////RgAAADQAAAAoAAAARU1GKypAAAAkAAAAGAAAAAAAgD8AAACAAAAAgAAAgD8AAACAAAAAgEYAAAAcAAAAEAAAAEVNRisCQAAADAAAAAAAAAAOAAAAFAAAAAAAAAAQAAAAFAAAAA==";

        private System.IO.Stream GetBinaryDataStream(string base64String)
        {
            return new System.IO.MemoryStream(System.Convert.FromBase64String(base64String));
        }

        #endregion

    }
}

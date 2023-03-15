using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml;
using A = DocumentFormat.OpenXml.Drawing;
using P14 = DocumentFormat.OpenXml.Office2010.PowerPoint;
using P15 = DocumentFormat.OpenXml.Office2013.PowerPoint;
using Thm15 = DocumentFormat.OpenXml.Office2013.Theme;
using A14 = DocumentFormat.OpenXml.Office2010.Drawing;

namespace GeneratedCode
{
    public class GeneratedClass
    {
        // Adds child parts and generates content of the specified part.
        public void CreateSlideMasterPart(SlideMasterPart part)
        {
            SlideLayoutPart slideLayoutPart1 = part.AddNewPart<SlideLayoutPart>("rId8");
            GenerateSlideLayoutPart1Content(slideLayoutPart1);

            SlideMasterPart slideMasterPart1 = slideLayoutPart1.AddNewPart<SlideMasterPart>("rId1");
            GenerateSlideMasterPart1Content(slideMasterPart1);

            slideMasterPart1.AddPart(slideLayoutPart1, "rId8");

            ThemePart themePart1 = slideMasterPart1.AddNewPart<ThemePart>("rId13");
            GenerateThemePart1Content(themePart1);

            SlideLayoutPart slideLayoutPart2 = slideMasterPart1.AddNewPart<SlideLayoutPart>("rId3");
            GenerateSlideLayoutPart2Content(slideLayoutPart2);

            slideLayoutPart2.AddPart(slideMasterPart1, "rId1");

            SlideLayoutPart slideLayoutPart3 = slideMasterPart1.AddNewPart<SlideLayoutPart>("rId7");
            GenerateSlideLayoutPart3Content(slideLayoutPart3);

            slideLayoutPart3.AddPart(slideMasterPart1, "rId1");

            SlideLayoutPart slideLayoutPart4 = slideMasterPart1.AddNewPart<SlideLayoutPart>("rId12");
            GenerateSlideLayoutPart4Content(slideLayoutPart4);

            ImagePart imagePart1 = slideLayoutPart4.AddNewPart<ImagePart>("image/jpeg", "rId2");
            GenerateImagePart1Content(imagePart1);

            slideLayoutPart4.AddPart(slideMasterPart1, "rId1");

            SlideLayoutPart slideLayoutPart5 = slideMasterPart1.AddNewPart<SlideLayoutPart>("rId2");
            GenerateSlideLayoutPart5Content(slideLayoutPart5);

            slideLayoutPart5.AddPart(slideMasterPart1, "rId1");

            SlideLayoutPart slideLayoutPart6 = slideMasterPart1.AddNewPart<SlideLayoutPart>("rId1");
            GenerateSlideLayoutPart6Content(slideLayoutPart6);

            slideLayoutPart6.AddPart(slideMasterPart1, "rId1");

            SlideLayoutPart slideLayoutPart7 = slideMasterPart1.AddNewPart<SlideLayoutPart>("rId6");
            GenerateSlideLayoutPart7Content(slideLayoutPart7);

            slideLayoutPart7.AddPart(slideMasterPart1, "rId1");

            SlideLayoutPart slideLayoutPart8 = slideMasterPart1.AddNewPart<SlideLayoutPart>("rId11");
            GenerateSlideLayoutPart8Content(slideLayoutPart8);

            slideLayoutPart8.AddPart(slideMasterPart1, "rId1");

            SlideLayoutPart slideLayoutPart9 = slideMasterPart1.AddNewPart<SlideLayoutPart>("rId5");
            GenerateSlideLayoutPart9Content(slideLayoutPart9);

            slideLayoutPart9.AddPart(slideMasterPart1, "rId1");

            SlideLayoutPart slideLayoutPart10 = slideMasterPart1.AddNewPart<SlideLayoutPart>("rId10");
            GenerateSlideLayoutPart10Content(slideLayoutPart10);

            slideLayoutPart10.AddPart(slideMasterPart1, "rId1");

            SlideLayoutPart slideLayoutPart11 = slideMasterPart1.AddNewPart<SlideLayoutPart>("rId4");
            GenerateSlideLayoutPart11Content(slideLayoutPart11);

            slideLayoutPart11.AddPart(slideMasterPart1, "rId1");

            SlideLayoutPart slideLayoutPart12 = slideMasterPart1.AddNewPart<SlideLayoutPart>("rId9");
            GenerateSlideLayoutPart12Content(slideLayoutPart12);

            slideLayoutPart12.AddPart(slideMasterPart1, "rId1");

            part.AddPart(themePart1, "rId13");

            part.AddPart(slideLayoutPart2, "rId3");

            part.AddPart(slideLayoutPart3, "rId7");

            part.AddPart(slideLayoutPart4, "rId12");

            part.AddPart(slideLayoutPart5, "rId2");

            part.AddPart(slideLayoutPart6, "rId1");

            part.AddPart(slideLayoutPart7, "rId6");

            part.AddPart(slideLayoutPart8, "rId11");

            part.AddPart(slideLayoutPart9, "rId5");

            part.AddPart(slideLayoutPart10, "rId10");

            part.AddPart(slideLayoutPart11, "rId4");

            part.AddPart(slideLayoutPart12, "rId9");

            GeneratePartContent(part);

        }

        // Generates content of slideLayoutPart1.
        private void GenerateSlideLayoutPart1Content(SlideLayoutPart slideLayoutPart1)
        {
            SlideLayout slideLayout1 = new SlideLayout(){ Type = SlideLayoutValues.ObjectText, Preserve = true };
            slideLayout1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            slideLayout1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            slideLayout1.AddNamespaceDeclaration("p", "http://schemas.openxmlformats.org/presentationml/2006/main");

            CommonSlideData commonSlideData1 = new CommonSlideData(){ Name = "Content with Caption" };

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

            Shape shape1 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties1 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties2 = new NonVisualDrawingProperties(){ Id = (UInt32Value)2U, Name = "Title 1" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties1 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks1 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties1.Append(shapeLocks1);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties2 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape1 = new PlaceholderShape(){ Type = PlaceholderValues.Title };

            applicationNonVisualDrawingProperties2.Append(placeholderShape1);

            nonVisualShapeProperties1.Append(nonVisualDrawingProperties2);
            nonVisualShapeProperties1.Append(nonVisualShapeDrawingProperties1);
            nonVisualShapeProperties1.Append(applicationNonVisualDrawingProperties2);

            ShapeProperties shapeProperties1 = new ShapeProperties();

            A.Transform2D transform2D1 = new A.Transform2D();
            A.Offset offset2 = new A.Offset(){ X = 524868L, Y = 381000L };
            A.Extents extents2 = new A.Extents(){ Cx = 2457648L, Cy = 1333500L };

            transform2D1.Append(offset2);
            transform2D1.Append(extents2);

            shapeProperties1.Append(transform2D1);

            TextBody textBody1 = new TextBody();
            A.BodyProperties bodyProperties1 = new A.BodyProperties(){ Anchor = A.TextAnchoringTypeValues.Bottom };

            A.ListStyle listStyle1 = new A.ListStyle();

            A.Level1ParagraphProperties level1ParagraphProperties1 = new A.Level1ParagraphProperties();
            A.DefaultRunProperties defaultRunProperties1 = new A.DefaultRunProperties(){ FontSize = 2667 };

            level1ParagraphProperties1.Append(defaultRunProperties1);

            listStyle1.Append(level1ParagraphProperties1);

            A.Paragraph paragraph1 = new A.Paragraph();

            A.Run run1 = new A.Run();
            A.RunProperties runProperties1 = new A.RunProperties(){ Language = "en-GB" };
            A.Text text1 = new A.Text();
            text1.Text = "Click to edit Master title style";

            run1.Append(runProperties1);
            run1.Append(text1);
            A.EndParagraphRunProperties endParagraphRunProperties1 = new A.EndParagraphRunProperties(){ Language = "en-US", Dirty = false };

            paragraph1.Append(run1);
            paragraph1.Append(endParagraphRunProperties1);

            textBody1.Append(bodyProperties1);
            textBody1.Append(listStyle1);
            textBody1.Append(paragraph1);

            shape1.Append(nonVisualShapeProperties1);
            shape1.Append(shapeProperties1);
            shape1.Append(textBody1);

            Shape shape2 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties2 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties3 = new NonVisualDrawingProperties(){ Id = (UInt32Value)3U, Name = "Content Placeholder 2" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties2 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks2 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties2.Append(shapeLocks2);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties3 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape2 = new PlaceholderShape(){ Index = (UInt32Value)1U };

            applicationNonVisualDrawingProperties3.Append(placeholderShape2);

            nonVisualShapeProperties2.Append(nonVisualDrawingProperties3);
            nonVisualShapeProperties2.Append(nonVisualShapeDrawingProperties2);
            nonVisualShapeProperties2.Append(applicationNonVisualDrawingProperties3);

            ShapeProperties shapeProperties2 = new ShapeProperties();

            A.Transform2D transform2D2 = new A.Transform2D();
            A.Offset offset3 = new A.Offset(){ X = 3239493L, Y = 822856L };
            A.Extents extents3 = new A.Extents(){ Cx = 3857625L, Cy = 4061354L };

            transform2D2.Append(offset3);
            transform2D2.Append(extents3);

            shapeProperties2.Append(transform2D2);

            TextBody textBody2 = new TextBody();
            A.BodyProperties bodyProperties2 = new A.BodyProperties();

            A.ListStyle listStyle2 = new A.ListStyle();

            A.Level1ParagraphProperties level1ParagraphProperties2 = new A.Level1ParagraphProperties();
            A.DefaultRunProperties defaultRunProperties2 = new A.DefaultRunProperties(){ FontSize = 2667 };

            level1ParagraphProperties2.Append(defaultRunProperties2);

            A.Level2ParagraphProperties level2ParagraphProperties1 = new A.Level2ParagraphProperties();
            A.DefaultRunProperties defaultRunProperties3 = new A.DefaultRunProperties(){ FontSize = 2333 };

            level2ParagraphProperties1.Append(defaultRunProperties3);

            A.Level3ParagraphProperties level3ParagraphProperties1 = new A.Level3ParagraphProperties();
            A.DefaultRunProperties defaultRunProperties4 = new A.DefaultRunProperties(){ FontSize = 2000 };

            level3ParagraphProperties1.Append(defaultRunProperties4);

            A.Level4ParagraphProperties level4ParagraphProperties1 = new A.Level4ParagraphProperties();
            A.DefaultRunProperties defaultRunProperties5 = new A.DefaultRunProperties(){ FontSize = 1667 };

            level4ParagraphProperties1.Append(defaultRunProperties5);

            A.Level5ParagraphProperties level5ParagraphProperties1 = new A.Level5ParagraphProperties();
            A.DefaultRunProperties defaultRunProperties6 = new A.DefaultRunProperties(){ FontSize = 1667 };

            level5ParagraphProperties1.Append(defaultRunProperties6);

            A.Level6ParagraphProperties level6ParagraphProperties1 = new A.Level6ParagraphProperties();
            A.DefaultRunProperties defaultRunProperties7 = new A.DefaultRunProperties(){ FontSize = 1667 };

            level6ParagraphProperties1.Append(defaultRunProperties7);

            A.Level7ParagraphProperties level7ParagraphProperties1 = new A.Level7ParagraphProperties();
            A.DefaultRunProperties defaultRunProperties8 = new A.DefaultRunProperties(){ FontSize = 1667 };

            level7ParagraphProperties1.Append(defaultRunProperties8);

            A.Level8ParagraphProperties level8ParagraphProperties1 = new A.Level8ParagraphProperties();
            A.DefaultRunProperties defaultRunProperties9 = new A.DefaultRunProperties(){ FontSize = 1667 };

            level8ParagraphProperties1.Append(defaultRunProperties9);

            A.Level9ParagraphProperties level9ParagraphProperties1 = new A.Level9ParagraphProperties();
            A.DefaultRunProperties defaultRunProperties10 = new A.DefaultRunProperties(){ FontSize = 1667 };

            level9ParagraphProperties1.Append(defaultRunProperties10);

            listStyle2.Append(level1ParagraphProperties2);
            listStyle2.Append(level2ParagraphProperties1);
            listStyle2.Append(level3ParagraphProperties1);
            listStyle2.Append(level4ParagraphProperties1);
            listStyle2.Append(level5ParagraphProperties1);
            listStyle2.Append(level6ParagraphProperties1);
            listStyle2.Append(level7ParagraphProperties1);
            listStyle2.Append(level8ParagraphProperties1);
            listStyle2.Append(level9ParagraphProperties1);

            A.Paragraph paragraph2 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties1 = new A.ParagraphProperties(){ Level = 0 };

            A.Run run2 = new A.Run();
            A.RunProperties runProperties2 = new A.RunProperties(){ Language = "en-GB" };
            A.Text text2 = new A.Text();
            text2.Text = "Click to edit Master text styles";

            run2.Append(runProperties2);
            run2.Append(text2);

            paragraph2.Append(paragraphProperties1);
            paragraph2.Append(run2);

            A.Paragraph paragraph3 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties2 = new A.ParagraphProperties(){ Level = 1 };

            A.Run run3 = new A.Run();
            A.RunProperties runProperties3 = new A.RunProperties(){ Language = "en-GB" };
            A.Text text3 = new A.Text();
            text3.Text = "Second level";

            run3.Append(runProperties3);
            run3.Append(text3);

            paragraph3.Append(paragraphProperties2);
            paragraph3.Append(run3);

            A.Paragraph paragraph4 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties3 = new A.ParagraphProperties(){ Level = 2 };

            A.Run run4 = new A.Run();
            A.RunProperties runProperties4 = new A.RunProperties(){ Language = "en-GB" };
            A.Text text4 = new A.Text();
            text4.Text = "Third level";

            run4.Append(runProperties4);
            run4.Append(text4);

            paragraph4.Append(paragraphProperties3);
            paragraph4.Append(run4);

            A.Paragraph paragraph5 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties4 = new A.ParagraphProperties(){ Level = 3 };

            A.Run run5 = new A.Run();
            A.RunProperties runProperties5 = new A.RunProperties(){ Language = "en-GB" };
            A.Text text5 = new A.Text();
            text5.Text = "Fourth level";

            run5.Append(runProperties5);
            run5.Append(text5);

            paragraph5.Append(paragraphProperties4);
            paragraph5.Append(run5);

            A.Paragraph paragraph6 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties5 = new A.ParagraphProperties(){ Level = 4 };

            A.Run run6 = new A.Run();
            A.RunProperties runProperties6 = new A.RunProperties(){ Language = "en-GB" };
            A.Text text6 = new A.Text();
            text6.Text = "Fifth level";

            run6.Append(runProperties6);
            run6.Append(text6);
            A.EndParagraphRunProperties endParagraphRunProperties2 = new A.EndParagraphRunProperties(){ Language = "en-US", Dirty = false };

            paragraph6.Append(paragraphProperties5);
            paragraph6.Append(run6);
            paragraph6.Append(endParagraphRunProperties2);

            textBody2.Append(bodyProperties2);
            textBody2.Append(listStyle2);
            textBody2.Append(paragraph2);
            textBody2.Append(paragraph3);
            textBody2.Append(paragraph4);
            textBody2.Append(paragraph5);
            textBody2.Append(paragraph6);

            shape2.Append(nonVisualShapeProperties2);
            shape2.Append(shapeProperties2);
            shape2.Append(textBody2);

            Shape shape3 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties3 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties4 = new NonVisualDrawingProperties(){ Id = (UInt32Value)4U, Name = "Text Placeholder 3" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties3 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks3 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties3.Append(shapeLocks3);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties4 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape3 = new PlaceholderShape(){ Type = PlaceholderValues.Body, Size = PlaceholderSizeValues.Half, Index = (UInt32Value)2U };

            applicationNonVisualDrawingProperties4.Append(placeholderShape3);

            nonVisualShapeProperties3.Append(nonVisualDrawingProperties4);
            nonVisualShapeProperties3.Append(nonVisualShapeDrawingProperties3);
            nonVisualShapeProperties3.Append(applicationNonVisualDrawingProperties4);

            ShapeProperties shapeProperties3 = new ShapeProperties();

            A.Transform2D transform2D3 = new A.Transform2D();
            A.Offset offset4 = new A.Offset(){ X = 524868L, Y = 1714500L };
            A.Extents extents4 = new A.Extents(){ Cx = 2457648L, Cy = 3176323L };

            transform2D3.Append(offset4);
            transform2D3.Append(extents4);

            shapeProperties3.Append(transform2D3);

            TextBody textBody3 = new TextBody();
            A.BodyProperties bodyProperties3 = new A.BodyProperties();

            A.ListStyle listStyle3 = new A.ListStyle();

            A.Level1ParagraphProperties level1ParagraphProperties3 = new A.Level1ParagraphProperties(){ LeftMargin = 0, Indent = 0 };
            A.NoBullet noBullet1 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties11 = new A.DefaultRunProperties(){ FontSize = 1333 };

            level1ParagraphProperties3.Append(noBullet1);
            level1ParagraphProperties3.Append(defaultRunProperties11);

            A.Level2ParagraphProperties level2ParagraphProperties2 = new A.Level2ParagraphProperties(){ LeftMargin = 380985, Indent = 0 };
            A.NoBullet noBullet2 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties12 = new A.DefaultRunProperties(){ FontSize = 1167 };

            level2ParagraphProperties2.Append(noBullet2);
            level2ParagraphProperties2.Append(defaultRunProperties12);

            A.Level3ParagraphProperties level3ParagraphProperties2 = new A.Level3ParagraphProperties(){ LeftMargin = 761970, Indent = 0 };
            A.NoBullet noBullet3 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties13 = new A.DefaultRunProperties(){ FontSize = 1000 };

            level3ParagraphProperties2.Append(noBullet3);
            level3ParagraphProperties2.Append(defaultRunProperties13);

            A.Level4ParagraphProperties level4ParagraphProperties2 = new A.Level4ParagraphProperties(){ LeftMargin = 1142954, Indent = 0 };
            A.NoBullet noBullet4 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties14 = new A.DefaultRunProperties(){ FontSize = 833 };

            level4ParagraphProperties2.Append(noBullet4);
            level4ParagraphProperties2.Append(defaultRunProperties14);

            A.Level5ParagraphProperties level5ParagraphProperties2 = new A.Level5ParagraphProperties(){ LeftMargin = 1523939, Indent = 0 };
            A.NoBullet noBullet5 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties15 = new A.DefaultRunProperties(){ FontSize = 833 };

            level5ParagraphProperties2.Append(noBullet5);
            level5ParagraphProperties2.Append(defaultRunProperties15);

            A.Level6ParagraphProperties level6ParagraphProperties2 = new A.Level6ParagraphProperties(){ LeftMargin = 1904924, Indent = 0 };
            A.NoBullet noBullet6 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties16 = new A.DefaultRunProperties(){ FontSize = 833 };

            level6ParagraphProperties2.Append(noBullet6);
            level6ParagraphProperties2.Append(defaultRunProperties16);

            A.Level7ParagraphProperties level7ParagraphProperties2 = new A.Level7ParagraphProperties(){ LeftMargin = 2285909, Indent = 0 };
            A.NoBullet noBullet7 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties17 = new A.DefaultRunProperties(){ FontSize = 833 };

            level7ParagraphProperties2.Append(noBullet7);
            level7ParagraphProperties2.Append(defaultRunProperties17);

            A.Level8ParagraphProperties level8ParagraphProperties2 = new A.Level8ParagraphProperties(){ LeftMargin = 2666893, Indent = 0 };
            A.NoBullet noBullet8 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties18 = new A.DefaultRunProperties(){ FontSize = 833 };

            level8ParagraphProperties2.Append(noBullet8);
            level8ParagraphProperties2.Append(defaultRunProperties18);

            A.Level9ParagraphProperties level9ParagraphProperties2 = new A.Level9ParagraphProperties(){ LeftMargin = 3047878, Indent = 0 };
            A.NoBullet noBullet9 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties19 = new A.DefaultRunProperties(){ FontSize = 833 };

            level9ParagraphProperties2.Append(noBullet9);
            level9ParagraphProperties2.Append(defaultRunProperties19);

            listStyle3.Append(level1ParagraphProperties3);
            listStyle3.Append(level2ParagraphProperties2);
            listStyle3.Append(level3ParagraphProperties2);
            listStyle3.Append(level4ParagraphProperties2);
            listStyle3.Append(level5ParagraphProperties2);
            listStyle3.Append(level6ParagraphProperties2);
            listStyle3.Append(level7ParagraphProperties2);
            listStyle3.Append(level8ParagraphProperties2);
            listStyle3.Append(level9ParagraphProperties2);

            A.Paragraph paragraph7 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties6 = new A.ParagraphProperties(){ Level = 0 };

            A.Run run7 = new A.Run();
            A.RunProperties runProperties7 = new A.RunProperties(){ Language = "en-GB" };
            A.Text text7 = new A.Text();
            text7.Text = "Click to edit Master text styles";

            run7.Append(runProperties7);
            run7.Append(text7);

            paragraph7.Append(paragraphProperties6);
            paragraph7.Append(run7);

            textBody3.Append(bodyProperties3);
            textBody3.Append(listStyle3);
            textBody3.Append(paragraph7);

            shape3.Append(nonVisualShapeProperties3);
            shape3.Append(shapeProperties3);
            shape3.Append(textBody3);

            Shape shape4 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties4 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties5 = new NonVisualDrawingProperties(){ Id = (UInt32Value)5U, Name = "Date Placeholder 4" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties4 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks4 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties4.Append(shapeLocks4);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties5 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape4 = new PlaceholderShape(){ Type = PlaceholderValues.DateAndTime, Size = PlaceholderSizeValues.Half, Index = (UInt32Value)10U };

            applicationNonVisualDrawingProperties5.Append(placeholderShape4);

            nonVisualShapeProperties4.Append(nonVisualDrawingProperties5);
            nonVisualShapeProperties4.Append(nonVisualShapeDrawingProperties4);
            nonVisualShapeProperties4.Append(applicationNonVisualDrawingProperties5);
            ShapeProperties shapeProperties4 = new ShapeProperties();

            TextBody textBody4 = new TextBody();
            A.BodyProperties bodyProperties4 = new A.BodyProperties();
            A.ListStyle listStyle4 = new A.ListStyle();

            A.Paragraph paragraph8 = new A.Paragraph();

            A.Field field1 = new A.Field(){ Id = "{54A533B2-3162-4708-BC91-D33C36DE81D4}", Type = "datetime1" };

            A.RunProperties runProperties8 = new A.RunProperties(){ Language = "en-US" };
            runProperties8.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            A.Text text8 = new A.Text();
            text8.Text = "1/17/2023";

            field1.Append(runProperties8);
            field1.Append(text8);
            A.EndParagraphRunProperties endParagraphRunProperties3 = new A.EndParagraphRunProperties(){ Language = "en-US", Dirty = false };

            paragraph8.Append(field1);
            paragraph8.Append(endParagraphRunProperties3);

            textBody4.Append(bodyProperties4);
            textBody4.Append(listStyle4);
            textBody4.Append(paragraph8);

            shape4.Append(nonVisualShapeProperties4);
            shape4.Append(shapeProperties4);
            shape4.Append(textBody4);

            Shape shape5 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties5 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties6 = new NonVisualDrawingProperties(){ Id = (UInt32Value)6U, Name = "Footer Placeholder 5" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties5 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks5 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties5.Append(shapeLocks5);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties6 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape5 = new PlaceholderShape(){ Type = PlaceholderValues.Footer, Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)11U };

            applicationNonVisualDrawingProperties6.Append(placeholderShape5);

            nonVisualShapeProperties5.Append(nonVisualDrawingProperties6);
            nonVisualShapeProperties5.Append(nonVisualShapeDrawingProperties5);
            nonVisualShapeProperties5.Append(applicationNonVisualDrawingProperties6);
            ShapeProperties shapeProperties5 = new ShapeProperties();

            TextBody textBody5 = new TextBody();
            A.BodyProperties bodyProperties5 = new A.BodyProperties();
            A.ListStyle listStyle5 = new A.ListStyle();

            A.Paragraph paragraph9 = new A.Paragraph();

            A.Run run8 = new A.Run();
            A.RunProperties runProperties9 = new A.RunProperties(){ Language = "en-US" };
            A.Text text9 = new A.Text();
            text9.Text = "Commercial & Workout Details";

            run8.Append(runProperties9);
            run8.Append(text9);
            A.EndParagraphRunProperties endParagraphRunProperties4 = new A.EndParagraphRunProperties(){ Language = "en-US", Dirty = false };

            paragraph9.Append(run8);
            paragraph9.Append(endParagraphRunProperties4);

            textBody5.Append(bodyProperties5);
            textBody5.Append(listStyle5);
            textBody5.Append(paragraph9);

            shape5.Append(nonVisualShapeProperties5);
            shape5.Append(shapeProperties5);
            shape5.Append(textBody5);

            Shape shape6 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties6 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties7 = new NonVisualDrawingProperties(){ Id = (UInt32Value)7U, Name = "Slide Number Placeholder 6" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties6 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks6 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties6.Append(shapeLocks6);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties7 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape6 = new PlaceholderShape(){ Type = PlaceholderValues.SlideNumber, Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)12U };

            applicationNonVisualDrawingProperties7.Append(placeholderShape6);

            nonVisualShapeProperties6.Append(nonVisualDrawingProperties7);
            nonVisualShapeProperties6.Append(nonVisualShapeDrawingProperties6);
            nonVisualShapeProperties6.Append(applicationNonVisualDrawingProperties7);
            ShapeProperties shapeProperties6 = new ShapeProperties();

            TextBody textBody6 = new TextBody();
            A.BodyProperties bodyProperties6 = new A.BodyProperties();
            A.ListStyle listStyle6 = new A.ListStyle();

            A.Paragraph paragraph10 = new A.Paragraph();

            A.Run run9 = new A.Run();

            A.RunProperties runProperties10 = new A.RunProperties(){ Language = "en-US" };

            A.SolidFill solidFill1 = new A.SolidFill();
            A.SchemeColor schemeColor1 = new A.SchemeColor(){ Val = A.SchemeColorValues.Accent2 };

            solidFill1.Append(schemeColor1);

            runProperties10.Append(solidFill1);
            A.Text text10 = new A.Text();
            text10.Text = "|";

            run9.Append(runProperties10);
            run9.Append(text10);

            A.Run run10 = new A.Run();
            A.RunProperties runProperties11 = new A.RunProperties(){ Language = "en-US" };
            A.Text text11 = new A.Text();
            text11.Text = "";

            run10.Append(runProperties11);
            run10.Append(text11);

            A.Field field2 = new A.Field(){ Id = "{E4F84C54-E2A4-46FF-B5B0-8F7A23C41D82}", Type = "slidenum" };

            A.RunProperties runProperties12 = new A.RunProperties(){ Language = "cs-CZ" };
            runProperties12.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            A.ParagraphProperties paragraphProperties7 = new A.ParagraphProperties();
            A.Text text12 = new A.Text();
            text12.Text = "‹#›";

            field2.Append(runProperties12);
            field2.Append(paragraphProperties7);
            field2.Append(text12);
            A.EndParagraphRunProperties endParagraphRunProperties5 = new A.EndParagraphRunProperties(){ Language = "cs-CZ", Dirty = false };

            paragraph10.Append(run9);
            paragraph10.Append(run10);
            paragraph10.Append(field2);
            paragraph10.Append(endParagraphRunProperties5);

            textBody6.Append(bodyProperties6);
            textBody6.Append(listStyle6);
            textBody6.Append(paragraph10);

            shape6.Append(nonVisualShapeProperties6);
            shape6.Append(shapeProperties6);
            shape6.Append(textBody6);

            shapeTree1.Append(nonVisualGroupShapeProperties1);
            shapeTree1.Append(groupShapeProperties1);
            shapeTree1.Append(shape1);
            shapeTree1.Append(shape2);
            shapeTree1.Append(shape3);
            shapeTree1.Append(shape4);
            shapeTree1.Append(shape5);
            shapeTree1.Append(shape6);

            CommonSlideDataExtensionList commonSlideDataExtensionList1 = new CommonSlideDataExtensionList();

            CommonSlideDataExtension commonSlideDataExtension1 = new CommonSlideDataExtension(){ Uri = "{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}" };

            P14.CreationId creationId1 = new P14.CreationId(){ Val = (UInt32Value)1362880519U };
            creationId1.AddNamespaceDeclaration("p14", "http://schemas.microsoft.com/office/powerpoint/2010/main");

            commonSlideDataExtension1.Append(creationId1);

            commonSlideDataExtensionList1.Append(commonSlideDataExtension1);

            commonSlideData1.Append(shapeTree1);
            commonSlideData1.Append(commonSlideDataExtensionList1);

            ColorMapOverride colorMapOverride1 = new ColorMapOverride();
            A.MasterColorMapping masterColorMapping1 = new A.MasterColorMapping();

            colorMapOverride1.Append(masterColorMapping1);

            slideLayout1.Append(commonSlideData1);
            slideLayout1.Append(colorMapOverride1);

            slideLayoutPart1.SlideLayout = slideLayout1;
        }

        // Generates content of slideMasterPart1.
        private void GenerateSlideMasterPart1Content(SlideMasterPart slideMasterPart1)
        {
            SlideMaster slideMaster1 = new SlideMaster();
            slideMaster1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            slideMaster1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            slideMaster1.AddNamespaceDeclaration("p", "http://schemas.openxmlformats.org/presentationml/2006/main");

            CommonSlideData commonSlideData2 = new CommonSlideData();

            Background background1 = new Background();

            BackgroundStyleReference backgroundStyleReference1 = new BackgroundStyleReference(){ Index = (UInt32Value)1001U };
            A.SchemeColor schemeColor2 = new A.SchemeColor(){ Val = A.SchemeColorValues.Background1 };

            backgroundStyleReference1.Append(schemeColor2);

            background1.Append(backgroundStyleReference1);

            ShapeTree shapeTree2 = new ShapeTree();

            NonVisualGroupShapeProperties nonVisualGroupShapeProperties2 = new NonVisualGroupShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties8 = new NonVisualDrawingProperties(){ Id = (UInt32Value)1U, Name = "" };
            NonVisualGroupShapeDrawingProperties nonVisualGroupShapeDrawingProperties2 = new NonVisualGroupShapeDrawingProperties();
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties8 = new ApplicationNonVisualDrawingProperties();

            nonVisualGroupShapeProperties2.Append(nonVisualDrawingProperties8);
            nonVisualGroupShapeProperties2.Append(nonVisualGroupShapeDrawingProperties2);
            nonVisualGroupShapeProperties2.Append(applicationNonVisualDrawingProperties8);

            GroupShapeProperties groupShapeProperties2 = new GroupShapeProperties();

            A.TransformGroup transformGroup2 = new A.TransformGroup();
            A.Offset offset5 = new A.Offset(){ X = 0L, Y = 0L };
            A.Extents extents5 = new A.Extents(){ Cx = 0L, Cy = 0L };
            A.ChildOffset childOffset2 = new A.ChildOffset(){ X = 0L, Y = 0L };
            A.ChildExtents childExtents2 = new A.ChildExtents(){ Cx = 0L, Cy = 0L };

            transformGroup2.Append(offset5);
            transformGroup2.Append(extents5);
            transformGroup2.Append(childOffset2);
            transformGroup2.Append(childExtents2);

            groupShapeProperties2.Append(transformGroup2);

            Shape shape7 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties7 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties9 = new NonVisualDrawingProperties(){ Id = (UInt32Value)2U, Name = "Title Placeholder 1" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties7 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks7 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties7.Append(shapeLocks7);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties9 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape7 = new PlaceholderShape(){ Type = PlaceholderValues.Title };

            applicationNonVisualDrawingProperties9.Append(placeholderShape7);

            nonVisualShapeProperties7.Append(nonVisualDrawingProperties9);
            nonVisualShapeProperties7.Append(nonVisualShapeDrawingProperties7);
            nonVisualShapeProperties7.Append(applicationNonVisualDrawingProperties9);

            ShapeProperties shapeProperties7 = new ShapeProperties();

            A.Transform2D transform2D4 = new A.Transform2D();
            A.Offset offset6 = new A.Offset(){ X = 523875L, Y = 304272L };
            A.Extents extents6 = new A.Extents(){ Cx = 6572250L, Cy = 1104636L };

            transform2D4.Append(offset6);
            transform2D4.Append(extents6);

            A.PresetGeometry presetGeometry1 = new A.PresetGeometry(){ Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList1 = new A.AdjustValueList();

            presetGeometry1.Append(adjustValueList1);

            shapeProperties7.Append(transform2D4);
            shapeProperties7.Append(presetGeometry1);

            TextBody textBody7 = new TextBody();

            A.BodyProperties bodyProperties7 = new A.BodyProperties(){ Vertical = A.TextVerticalValues.Horizontal, LeftInset = 91440, TopInset = 45720, RightInset = 91440, BottomInset = 45720, RightToLeftColumns = false, Anchor = A.TextAnchoringTypeValues.Center };
            A.NormalAutoFit normalAutoFit1 = new A.NormalAutoFit();

            bodyProperties7.Append(normalAutoFit1);
            A.ListStyle listStyle7 = new A.ListStyle();

            A.Paragraph paragraph11 = new A.Paragraph();

            A.Run run11 = new A.Run();
            A.RunProperties runProperties13 = new A.RunProperties(){ Language = "en-GB" };
            A.Text text13 = new A.Text();
            text13.Text = "Click to edit Master title style";

            run11.Append(runProperties13);
            run11.Append(text13);
            A.EndParagraphRunProperties endParagraphRunProperties6 = new A.EndParagraphRunProperties(){ Language = "en-US", Dirty = false };

            paragraph11.Append(run11);
            paragraph11.Append(endParagraphRunProperties6);

            textBody7.Append(bodyProperties7);
            textBody7.Append(listStyle7);
            textBody7.Append(paragraph11);

            shape7.Append(nonVisualShapeProperties7);
            shape7.Append(shapeProperties7);
            shape7.Append(textBody7);

            Shape shape8 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties8 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties10 = new NonVisualDrawingProperties(){ Id = (UInt32Value)3U, Name = "Text Placeholder 2" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties8 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks8 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties8.Append(shapeLocks8);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties10 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape8 = new PlaceholderShape(){ Type = PlaceholderValues.Body, Index = (UInt32Value)1U };

            applicationNonVisualDrawingProperties10.Append(placeholderShape8);

            nonVisualShapeProperties8.Append(nonVisualDrawingProperties10);
            nonVisualShapeProperties8.Append(nonVisualShapeDrawingProperties8);
            nonVisualShapeProperties8.Append(applicationNonVisualDrawingProperties10);

            ShapeProperties shapeProperties8 = new ShapeProperties();

            A.Transform2D transform2D5 = new A.Transform2D();
            A.Offset offset7 = new A.Offset(){ X = 523875L, Y = 1521354L };
            A.Extents extents7 = new A.Extents(){ Cx = 6572250L, Cy = 3626115L };

            transform2D5.Append(offset7);
            transform2D5.Append(extents7);

            A.PresetGeometry presetGeometry2 = new A.PresetGeometry(){ Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList2 = new A.AdjustValueList();

            presetGeometry2.Append(adjustValueList2);

            shapeProperties8.Append(transform2D5);
            shapeProperties8.Append(presetGeometry2);

            TextBody textBody8 = new TextBody();

            A.BodyProperties bodyProperties8 = new A.BodyProperties(){ Vertical = A.TextVerticalValues.Horizontal, LeftInset = 91440, TopInset = 45720, RightInset = 91440, BottomInset = 45720, RightToLeftColumns = false };
            A.NormalAutoFit normalAutoFit2 = new A.NormalAutoFit();

            bodyProperties8.Append(normalAutoFit2);
            A.ListStyle listStyle8 = new A.ListStyle();

            A.Paragraph paragraph12 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties8 = new A.ParagraphProperties(){ Level = 0 };

            A.Run run12 = new A.Run();
            A.RunProperties runProperties14 = new A.RunProperties(){ Language = "en-GB" };
            A.Text text14 = new A.Text();
            text14.Text = "Click to edit Master text styles";

            run12.Append(runProperties14);
            run12.Append(text14);

            paragraph12.Append(paragraphProperties8);
            paragraph12.Append(run12);

            A.Paragraph paragraph13 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties9 = new A.ParagraphProperties(){ Level = 1 };

            A.Run run13 = new A.Run();
            A.RunProperties runProperties15 = new A.RunProperties(){ Language = "en-GB" };
            A.Text text15 = new A.Text();
            text15.Text = "Second level";

            run13.Append(runProperties15);
            run13.Append(text15);

            paragraph13.Append(paragraphProperties9);
            paragraph13.Append(run13);

            A.Paragraph paragraph14 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties10 = new A.ParagraphProperties(){ Level = 2 };

            A.Run run14 = new A.Run();
            A.RunProperties runProperties16 = new A.RunProperties(){ Language = "en-GB" };
            A.Text text16 = new A.Text();
            text16.Text = "Third level";

            run14.Append(runProperties16);
            run14.Append(text16);

            paragraph14.Append(paragraphProperties10);
            paragraph14.Append(run14);

            A.Paragraph paragraph15 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties11 = new A.ParagraphProperties(){ Level = 3 };

            A.Run run15 = new A.Run();
            A.RunProperties runProperties17 = new A.RunProperties(){ Language = "en-GB" };
            A.Text text17 = new A.Text();
            text17.Text = "Fourth level";

            run15.Append(runProperties17);
            run15.Append(text17);

            paragraph15.Append(paragraphProperties11);
            paragraph15.Append(run15);

            A.Paragraph paragraph16 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties12 = new A.ParagraphProperties(){ Level = 4 };

            A.Run run16 = new A.Run();
            A.RunProperties runProperties18 = new A.RunProperties(){ Language = "en-GB" };
            A.Text text18 = new A.Text();
            text18.Text = "Fifth level";

            run16.Append(runProperties18);
            run16.Append(text18);
            A.EndParagraphRunProperties endParagraphRunProperties7 = new A.EndParagraphRunProperties(){ Language = "en-US", Dirty = false };

            paragraph16.Append(paragraphProperties12);
            paragraph16.Append(run16);
            paragraph16.Append(endParagraphRunProperties7);

            textBody8.Append(bodyProperties8);
            textBody8.Append(listStyle8);
            textBody8.Append(paragraph12);
            textBody8.Append(paragraph13);
            textBody8.Append(paragraph14);
            textBody8.Append(paragraph15);
            textBody8.Append(paragraph16);

            shape8.Append(nonVisualShapeProperties8);
            shape8.Append(shapeProperties8);
            shape8.Append(textBody8);

            Shape shape9 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties9 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties11 = new NonVisualDrawingProperties(){ Id = (UInt32Value)4U, Name = "Date Placeholder 3" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties9 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks9 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties9.Append(shapeLocks9);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties11 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape9 = new PlaceholderShape(){ Type = PlaceholderValues.DateAndTime, Size = PlaceholderSizeValues.Half, Index = (UInt32Value)2U };

            applicationNonVisualDrawingProperties11.Append(placeholderShape9);

            nonVisualShapeProperties9.Append(nonVisualDrawingProperties11);
            nonVisualShapeProperties9.Append(nonVisualShapeDrawingProperties9);
            nonVisualShapeProperties9.Append(applicationNonVisualDrawingProperties11);

            ShapeProperties shapeProperties9 = new ShapeProperties();

            A.Transform2D transform2D6 = new A.Transform2D();
            A.Offset offset8 = new A.Offset(){ X = 523875L, Y = 5296960L };
            A.Extents extents8 = new A.Extents(){ Cx = 1714500L, Cy = 304271L };

            transform2D6.Append(offset8);
            transform2D6.Append(extents8);

            A.PresetGeometry presetGeometry3 = new A.PresetGeometry(){ Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList3 = new A.AdjustValueList();

            presetGeometry3.Append(adjustValueList3);

            shapeProperties9.Append(transform2D6);
            shapeProperties9.Append(presetGeometry3);

            TextBody textBody9 = new TextBody();
            A.BodyProperties bodyProperties9 = new A.BodyProperties(){ Vertical = A.TextVerticalValues.Horizontal, LeftInset = 91440, TopInset = 45720, RightInset = 91440, BottomInset = 45720, RightToLeftColumns = false, Anchor = A.TextAnchoringTypeValues.Center };

            A.ListStyle listStyle9 = new A.ListStyle();

            A.Level1ParagraphProperties level1ParagraphProperties4 = new A.Level1ParagraphProperties(){ Alignment = A.TextAlignmentTypeValues.Left };

            A.DefaultRunProperties defaultRunProperties20 = new A.DefaultRunProperties(){ FontSize = 1000 };

            A.SolidFill solidFill2 = new A.SolidFill();

            A.SchemeColor schemeColor3 = new A.SchemeColor(){ Val = A.SchemeColorValues.Text1 };
            A.Tint tint1 = new A.Tint(){ Val = 75000 };

            schemeColor3.Append(tint1);

            solidFill2.Append(schemeColor3);

            defaultRunProperties20.Append(solidFill2);

            level1ParagraphProperties4.Append(defaultRunProperties20);

            listStyle9.Append(level1ParagraphProperties4);

            A.Paragraph paragraph17 = new A.Paragraph();

            A.Field field3 = new A.Field(){ Id = "{C5F68010-134A-445E-BD29-BEAAA1D9860C}", Type = "datetime1" };

            A.RunProperties runProperties19 = new A.RunProperties(){ Language = "en-US" };
            runProperties19.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            A.Text text19 = new A.Text();
            text19.Text = "1/17/2023";

            field3.Append(runProperties19);
            field3.Append(text19);
            A.EndParagraphRunProperties endParagraphRunProperties8 = new A.EndParagraphRunProperties(){ Language = "en-US", Dirty = false };

            paragraph17.Append(field3);
            paragraph17.Append(endParagraphRunProperties8);

            textBody9.Append(bodyProperties9);
            textBody9.Append(listStyle9);
            textBody9.Append(paragraph17);

            shape9.Append(nonVisualShapeProperties9);
            shape9.Append(shapeProperties9);
            shape9.Append(textBody9);

            Shape shape10 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties10 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties12 = new NonVisualDrawingProperties(){ Id = (UInt32Value)5U, Name = "Footer Placeholder 4" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties10 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks10 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties10.Append(shapeLocks10);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties12 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape10 = new PlaceholderShape(){ Type = PlaceholderValues.Footer, Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)3U };

            applicationNonVisualDrawingProperties12.Append(placeholderShape10);

            nonVisualShapeProperties10.Append(nonVisualDrawingProperties12);
            nonVisualShapeProperties10.Append(nonVisualShapeDrawingProperties10);
            nonVisualShapeProperties10.Append(applicationNonVisualDrawingProperties12);

            ShapeProperties shapeProperties10 = new ShapeProperties();

            A.Transform2D transform2D7 = new A.Transform2D();
            A.Offset offset9 = new A.Offset(){ X = 2524125L, Y = 5296960L };
            A.Extents extents9 = new A.Extents(){ Cx = 2571750L, Cy = 304271L };

            transform2D7.Append(offset9);
            transform2D7.Append(extents9);

            A.PresetGeometry presetGeometry4 = new A.PresetGeometry(){ Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList4 = new A.AdjustValueList();

            presetGeometry4.Append(adjustValueList4);

            shapeProperties10.Append(transform2D7);
            shapeProperties10.Append(presetGeometry4);

            TextBody textBody10 = new TextBody();
            A.BodyProperties bodyProperties10 = new A.BodyProperties(){ Vertical = A.TextVerticalValues.Horizontal, LeftInset = 91440, TopInset = 45720, RightInset = 91440, BottomInset = 45720, RightToLeftColumns = false, Anchor = A.TextAnchoringTypeValues.Center };

            A.ListStyle listStyle10 = new A.ListStyle();

            A.Level1ParagraphProperties level1ParagraphProperties5 = new A.Level1ParagraphProperties(){ Alignment = A.TextAlignmentTypeValues.Center };

            A.DefaultRunProperties defaultRunProperties21 = new A.DefaultRunProperties(){ FontSize = 1000 };

            A.SolidFill solidFill3 = new A.SolidFill();

            A.SchemeColor schemeColor4 = new A.SchemeColor(){ Val = A.SchemeColorValues.Text1 };
            A.Tint tint2 = new A.Tint(){ Val = 75000 };

            schemeColor4.Append(tint2);

            solidFill3.Append(schemeColor4);

            defaultRunProperties21.Append(solidFill3);

            level1ParagraphProperties5.Append(defaultRunProperties21);

            listStyle10.Append(level1ParagraphProperties5);

            A.Paragraph paragraph18 = new A.Paragraph();

            A.Run run17 = new A.Run();
            A.RunProperties runProperties20 = new A.RunProperties(){ Language = "en-US" };
            A.Text text20 = new A.Text();
            text20.Text = "Commercial & Workout Details";

            run17.Append(runProperties20);
            run17.Append(text20);
            A.EndParagraphRunProperties endParagraphRunProperties9 = new A.EndParagraphRunProperties(){ Language = "en-US", Dirty = false };

            paragraph18.Append(run17);
            paragraph18.Append(endParagraphRunProperties9);

            textBody10.Append(bodyProperties10);
            textBody10.Append(listStyle10);
            textBody10.Append(paragraph18);

            shape10.Append(nonVisualShapeProperties10);
            shape10.Append(shapeProperties10);
            shape10.Append(textBody10);

            Shape shape11 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties11 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties13 = new NonVisualDrawingProperties(){ Id = (UInt32Value)6U, Name = "Slide Number Placeholder 5" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties11 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks11 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties11.Append(shapeLocks11);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties13 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape11 = new PlaceholderShape(){ Type = PlaceholderValues.SlideNumber, Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)4U };

            applicationNonVisualDrawingProperties13.Append(placeholderShape11);

            nonVisualShapeProperties11.Append(nonVisualDrawingProperties13);
            nonVisualShapeProperties11.Append(nonVisualShapeDrawingProperties11);
            nonVisualShapeProperties11.Append(applicationNonVisualDrawingProperties13);

            ShapeProperties shapeProperties11 = new ShapeProperties();

            A.Transform2D transform2D8 = new A.Transform2D();
            A.Offset offset10 = new A.Offset(){ X = 5381625L, Y = 5296960L };
            A.Extents extents10 = new A.Extents(){ Cx = 1714500L, Cy = 304271L };

            transform2D8.Append(offset10);
            transform2D8.Append(extents10);

            A.PresetGeometry presetGeometry5 = new A.PresetGeometry(){ Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList5 = new A.AdjustValueList();

            presetGeometry5.Append(adjustValueList5);

            shapeProperties11.Append(transform2D8);
            shapeProperties11.Append(presetGeometry5);

            TextBody textBody11 = new TextBody();
            A.BodyProperties bodyProperties11 = new A.BodyProperties(){ Vertical = A.TextVerticalValues.Horizontal, LeftInset = 91440, TopInset = 45720, RightInset = 91440, BottomInset = 45720, RightToLeftColumns = false, Anchor = A.TextAnchoringTypeValues.Center };

            A.ListStyle listStyle11 = new A.ListStyle();

            A.Level1ParagraphProperties level1ParagraphProperties6 = new A.Level1ParagraphProperties(){ Alignment = A.TextAlignmentTypeValues.Right };

            A.DefaultRunProperties defaultRunProperties22 = new A.DefaultRunProperties(){ FontSize = 1000 };

            A.SolidFill solidFill4 = new A.SolidFill();

            A.SchemeColor schemeColor5 = new A.SchemeColor(){ Val = A.SchemeColorValues.Text1 };
            A.Tint tint3 = new A.Tint(){ Val = 75000 };

            schemeColor5.Append(tint3);

            solidFill4.Append(schemeColor5);

            defaultRunProperties22.Append(solidFill4);

            level1ParagraphProperties6.Append(defaultRunProperties22);

            listStyle11.Append(level1ParagraphProperties6);

            A.Paragraph paragraph19 = new A.Paragraph();

            A.Run run18 = new A.Run();

            A.RunProperties runProperties21 = new A.RunProperties(){ Language = "en-US" };

            A.SolidFill solidFill5 = new A.SolidFill();
            A.SchemeColor schemeColor6 = new A.SchemeColor(){ Val = A.SchemeColorValues.Accent2 };

            solidFill5.Append(schemeColor6);

            runProperties21.Append(solidFill5);
            A.Text text21 = new A.Text();
            text21.Text = "|";

            run18.Append(runProperties21);
            run18.Append(text21);

            A.Run run19 = new A.Run();
            A.RunProperties runProperties22 = new A.RunProperties(){ Language = "en-US" };
            A.Text text22 = new A.Text();
            text22.Text = "";

            run19.Append(runProperties22);
            run19.Append(text22);

            A.Field field4 = new A.Field(){ Id = "{E4F84C54-E2A4-46FF-B5B0-8F7A23C41D82}", Type = "slidenum" };

            A.RunProperties runProperties23 = new A.RunProperties(){ Language = "cs-CZ" };
            runProperties23.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            A.ParagraphProperties paragraphProperties13 = new A.ParagraphProperties();
            A.Text text23 = new A.Text();
            text23.Text = "‹#›";

            field4.Append(runProperties23);
            field4.Append(paragraphProperties13);
            field4.Append(text23);
            A.EndParagraphRunProperties endParagraphRunProperties10 = new A.EndParagraphRunProperties(){ Language = "cs-CZ", Dirty = false };

            paragraph19.Append(run18);
            paragraph19.Append(run19);
            paragraph19.Append(field4);
            paragraph19.Append(endParagraphRunProperties10);

            textBody11.Append(bodyProperties11);
            textBody11.Append(listStyle11);
            textBody11.Append(paragraph19);

            shape11.Append(nonVisualShapeProperties11);
            shape11.Append(shapeProperties11);
            shape11.Append(textBody11);

            Shape shape12 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties12 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties14 = new NonVisualDrawingProperties(){ Id = (UInt32Value)7U, Name = "Rectangle 6" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList1 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension1 = new A.NonVisualDrawingPropertiesExtension(){ Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            OpenXmlUnknownElement openXmlUnknownElement1 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{DABD3684-197E-4B2B-99DA-0FFE27DA5C10}\" />");

            nonVisualDrawingPropertiesExtension1.Append(openXmlUnknownElement1);

            nonVisualDrawingPropertiesExtensionList1.Append(nonVisualDrawingPropertiesExtension1);

            nonVisualDrawingProperties14.Append(nonVisualDrawingPropertiesExtensionList1);
            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties12 = new NonVisualShapeDrawingProperties();
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties14 = new ApplicationNonVisualDrawingProperties(){ UserDrawn = true };

            nonVisualShapeProperties12.Append(nonVisualDrawingProperties14);
            nonVisualShapeProperties12.Append(nonVisualShapeDrawingProperties12);
            nonVisualShapeProperties12.Append(applicationNonVisualDrawingProperties14);

            ShapeProperties shapeProperties12 = new ShapeProperties();

            A.Transform2D transform2D9 = new A.Transform2D();
            A.Offset offset11 = new A.Offset(){ X = 269874L, Y = 1075788L };
            A.Extents extents11 = new A.Extents(){ Cx = 7110678L, Cy = 36000L };

            transform2D9.Append(offset11);
            transform2D9.Append(extents11);

            A.PresetGeometry presetGeometry6 = new A.PresetGeometry(){ Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList6 = new A.AdjustValueList();

            presetGeometry6.Append(adjustValueList6);

            A.GradientFill gradientFill1 = new A.GradientFill(){ Flip = A.TileFlipValues.None, RotateWithShape = true };

            A.GradientStopList gradientStopList1 = new A.GradientStopList();

            A.GradientStop gradientStop1 = new A.GradientStop(){ Position = 0 };
            A.SchemeColor schemeColor7 = new A.SchemeColor(){ Val = A.SchemeColorValues.Accent2 };

            gradientStop1.Append(schemeColor7);

            A.GradientStop gradientStop2 = new A.GradientStop(){ Position = 100000 };
            A.SchemeColor schemeColor8 = new A.SchemeColor(){ Val = A.SchemeColorValues.Accent1 };

            gradientStop2.Append(schemeColor8);

            gradientStopList1.Append(gradientStop1);
            gradientStopList1.Append(gradientStop2);
            A.LinearGradientFill linearGradientFill1 = new A.LinearGradientFill(){ Angle = 0, Scaled = true };
            A.TileRectangle tileRectangle1 = new A.TileRectangle();

            gradientFill1.Append(gradientStopList1);
            gradientFill1.Append(linearGradientFill1);
            gradientFill1.Append(tileRectangle1);

            A.Outline outline1 = new A.Outline();
            A.NoFill noFill1 = new A.NoFill();

            outline1.Append(noFill1);

            shapeProperties12.Append(transform2D9);
            shapeProperties12.Append(presetGeometry6);
            shapeProperties12.Append(gradientFill1);
            shapeProperties12.Append(outline1);

            ShapeStyle shapeStyle1 = new ShapeStyle();

            A.LineReference lineReference1 = new A.LineReference(){ Index = (UInt32Value)2U };

            A.SchemeColor schemeColor9 = new A.SchemeColor(){ Val = A.SchemeColorValues.Accent1 };
            A.Shade shade1 = new A.Shade(){ Val = 50000 };

            schemeColor9.Append(shade1);

            lineReference1.Append(schemeColor9);

            A.FillReference fillReference1 = new A.FillReference(){ Index = (UInt32Value)1U };
            A.SchemeColor schemeColor10 = new A.SchemeColor(){ Val = A.SchemeColorValues.Accent1 };

            fillReference1.Append(schemeColor10);

            A.EffectReference effectReference1 = new A.EffectReference(){ Index = (UInt32Value)0U };
            A.SchemeColor schemeColor11 = new A.SchemeColor(){ Val = A.SchemeColorValues.Accent1 };

            effectReference1.Append(schemeColor11);

            A.FontReference fontReference1 = new A.FontReference(){ Index = A.FontCollectionIndexValues.Minor };
            A.SchemeColor schemeColor12 = new A.SchemeColor(){ Val = A.SchemeColorValues.Light1 };

            fontReference1.Append(schemeColor12);

            shapeStyle1.Append(lineReference1);
            shapeStyle1.Append(fillReference1);
            shapeStyle1.Append(effectReference1);
            shapeStyle1.Append(fontReference1);

            TextBody textBody12 = new TextBody();
            A.BodyProperties bodyProperties12 = new A.BodyProperties(){ RightToLeftColumns = false, Anchor = A.TextAnchoringTypeValues.Center };
            A.ListStyle listStyle12 = new A.ListStyle();

            A.Paragraph paragraph20 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties14 = new A.ParagraphProperties(){ Alignment = A.TextAlignmentTypeValues.Center };
            A.EndParagraphRunProperties endParagraphRunProperties11 = new A.EndParagraphRunProperties(){ Language = "cs-CZ", FontSize = 1404 };

            paragraph20.Append(paragraphProperties14);
            paragraph20.Append(endParagraphRunProperties11);

            textBody12.Append(bodyProperties12);
            textBody12.Append(listStyle12);
            textBody12.Append(paragraph20);

            shape12.Append(nonVisualShapeProperties12);
            shape12.Append(shapeProperties12);
            shape12.Append(shapeStyle1);
            shape12.Append(textBody12);

            shapeTree2.Append(nonVisualGroupShapeProperties2);
            shapeTree2.Append(groupShapeProperties2);
            shapeTree2.Append(shape7);
            shapeTree2.Append(shape8);
            shapeTree2.Append(shape9);
            shapeTree2.Append(shape10);
            shapeTree2.Append(shape11);
            shapeTree2.Append(shape12);

            CommonSlideDataExtensionList commonSlideDataExtensionList2 = new CommonSlideDataExtensionList();

            CommonSlideDataExtension commonSlideDataExtension2 = new CommonSlideDataExtension(){ Uri = "{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}" };

            P14.CreationId creationId2 = new P14.CreationId(){ Val = (UInt32Value)1735501897U };
            creationId2.AddNamespaceDeclaration("p14", "http://schemas.microsoft.com/office/powerpoint/2010/main");

            commonSlideDataExtension2.Append(creationId2);

            commonSlideDataExtensionList2.Append(commonSlideDataExtension2);

            commonSlideData2.Append(background1);
            commonSlideData2.Append(shapeTree2);
            commonSlideData2.Append(commonSlideDataExtensionList2);
            ColorMap colorMap1 = new ColorMap(){ Background1 = A.ColorSchemeIndexValues.Light1, Text1 = A.ColorSchemeIndexValues.Dark1, Background2 = A.ColorSchemeIndexValues.Light2, Text2 = A.ColorSchemeIndexValues.Dark2, Accent1 = A.ColorSchemeIndexValues.Accent1, Accent2 = A.ColorSchemeIndexValues.Accent2, Accent3 = A.ColorSchemeIndexValues.Accent3, Accent4 = A.ColorSchemeIndexValues.Accent4, Accent5 = A.ColorSchemeIndexValues.Accent5, Accent6 = A.ColorSchemeIndexValues.Accent6, Hyperlink = A.ColorSchemeIndexValues.Hyperlink, FollowedHyperlink = A.ColorSchemeIndexValues.FollowedHyperlink };

            SlideLayoutIdList slideLayoutIdList1 = new SlideLayoutIdList();
            SlideLayoutId slideLayoutId1 = new SlideLayoutId(){ Id = (UInt32Value)2147483663U, RelationshipId = "rId1" };
            SlideLayoutId slideLayoutId2 = new SlideLayoutId(){ Id = (UInt32Value)2147483664U, RelationshipId = "rId2" };
            SlideLayoutId slideLayoutId3 = new SlideLayoutId(){ Id = (UInt32Value)2147483665U, RelationshipId = "rId3" };
            SlideLayoutId slideLayoutId4 = new SlideLayoutId(){ Id = (UInt32Value)2147483666U, RelationshipId = "rId4" };
            SlideLayoutId slideLayoutId5 = new SlideLayoutId(){ Id = (UInt32Value)2147483667U, RelationshipId = "rId5" };
            SlideLayoutId slideLayoutId6 = new SlideLayoutId(){ Id = (UInt32Value)2147483668U, RelationshipId = "rId6" };
            SlideLayoutId slideLayoutId7 = new SlideLayoutId(){ Id = (UInt32Value)2147483669U, RelationshipId = "rId7" };
            SlideLayoutId slideLayoutId8 = new SlideLayoutId(){ Id = (UInt32Value)2147483670U, RelationshipId = "rId8" };
            SlideLayoutId slideLayoutId9 = new SlideLayoutId(){ Id = (UInt32Value)2147483671U, RelationshipId = "rId9" };
            SlideLayoutId slideLayoutId10 = new SlideLayoutId(){ Id = (UInt32Value)2147483672U, RelationshipId = "rId10" };
            SlideLayoutId slideLayoutId11 = new SlideLayoutId(){ Id = (UInt32Value)2147483673U, RelationshipId = "rId11" };
            SlideLayoutId slideLayoutId12 = new SlideLayoutId(){ Id = (UInt32Value)2147483676U, RelationshipId = "rId12" };

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
            slideLayoutIdList1.Append(slideLayoutId12);
            HeaderFooter headerFooter1 = new HeaderFooter(){ Header = false, DateTime = false };

            TextStyles textStyles1 = new TextStyles();

            TitleStyle titleStyle1 = new TitleStyle();

            A.Level1ParagraphProperties level1ParagraphProperties7 = new A.Level1ParagraphProperties(){ Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 761970, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.LineSpacing lineSpacing1 = new A.LineSpacing();
            A.SpacingPercent spacingPercent1 = new A.SpacingPercent(){ Val = 90000 };

            lineSpacing1.Append(spacingPercent1);

            A.SpaceBefore spaceBefore1 = new A.SpaceBefore();
            A.SpacingPercent spacingPercent2 = new A.SpacingPercent(){ Val = 0 };

            spaceBefore1.Append(spacingPercent2);
            A.NoBullet noBullet10 = new A.NoBullet();

            A.DefaultRunProperties defaultRunProperties23 = new A.DefaultRunProperties(){ FontSize = 3667, Kerning = 1200 };

            A.SolidFill solidFill6 = new A.SolidFill();
            A.SchemeColor schemeColor13 = new A.SchemeColor(){ Val = A.SchemeColorValues.Text1 };

            solidFill6.Append(schemeColor13);
            A.LatinFont latinFont1 = new A.LatinFont(){ Typeface = "+mj-lt" };
            A.EastAsianFont eastAsianFont1 = new A.EastAsianFont(){ Typeface = "+mj-ea" };
            A.ComplexScriptFont complexScriptFont1 = new A.ComplexScriptFont(){ Typeface = "+mj-cs" };

            defaultRunProperties23.Append(solidFill6);
            defaultRunProperties23.Append(latinFont1);
            defaultRunProperties23.Append(eastAsianFont1);
            defaultRunProperties23.Append(complexScriptFont1);

            level1ParagraphProperties7.Append(lineSpacing1);
            level1ParagraphProperties7.Append(spaceBefore1);
            level1ParagraphProperties7.Append(noBullet10);
            level1ParagraphProperties7.Append(defaultRunProperties23);

            titleStyle1.Append(level1ParagraphProperties7);

            BodyStyle bodyStyle1 = new BodyStyle();

            A.Level1ParagraphProperties level1ParagraphProperties8 = new A.Level1ParagraphProperties(){ LeftMargin = 190492, Indent = -190492, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 761970, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.LineSpacing lineSpacing2 = new A.LineSpacing();
            A.SpacingPercent spacingPercent3 = new A.SpacingPercent(){ Val = 90000 };

            lineSpacing2.Append(spacingPercent3);

            A.SpaceBefore spaceBefore2 = new A.SpaceBefore();
            A.SpacingPoints spacingPoints1 = new A.SpacingPoints(){ Val = 833 };

            spaceBefore2.Append(spacingPoints1);
            A.BulletFont bulletFont1 = new A.BulletFont(){ Typeface = "Arial", Panose = "020B0604020202020204", PitchFamily = 34, CharacterSet = 0 };
            A.CharacterBullet characterBullet1 = new A.CharacterBullet(){ Char = "•" };

            A.DefaultRunProperties defaultRunProperties24 = new A.DefaultRunProperties(){ FontSize = 2333, Kerning = 1200 };

            A.SolidFill solidFill7 = new A.SolidFill();
            A.SchemeColor schemeColor14 = new A.SchemeColor(){ Val = A.SchemeColorValues.Text1 };

            solidFill7.Append(schemeColor14);
            A.LatinFont latinFont2 = new A.LatinFont(){ Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont2 = new A.EastAsianFont(){ Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont2 = new A.ComplexScriptFont(){ Typeface = "+mn-cs" };

            defaultRunProperties24.Append(solidFill7);
            defaultRunProperties24.Append(latinFont2);
            defaultRunProperties24.Append(eastAsianFont2);
            defaultRunProperties24.Append(complexScriptFont2);

            level1ParagraphProperties8.Append(lineSpacing2);
            level1ParagraphProperties8.Append(spaceBefore2);
            level1ParagraphProperties8.Append(bulletFont1);
            level1ParagraphProperties8.Append(characterBullet1);
            level1ParagraphProperties8.Append(defaultRunProperties24);

            A.Level2ParagraphProperties level2ParagraphProperties3 = new A.Level2ParagraphProperties(){ LeftMargin = 571477, Indent = -190492, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 761970, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.LineSpacing lineSpacing3 = new A.LineSpacing();
            A.SpacingPercent spacingPercent4 = new A.SpacingPercent(){ Val = 90000 };

            lineSpacing3.Append(spacingPercent4);

            A.SpaceBefore spaceBefore3 = new A.SpaceBefore();
            A.SpacingPoints spacingPoints2 = new A.SpacingPoints(){ Val = 417 };

            spaceBefore3.Append(spacingPoints2);
            A.BulletFont bulletFont2 = new A.BulletFont(){ Typeface = "Arial", Panose = "020B0604020202020204", PitchFamily = 34, CharacterSet = 0 };
            A.CharacterBullet characterBullet2 = new A.CharacterBullet(){ Char = "•" };

            A.DefaultRunProperties defaultRunProperties25 = new A.DefaultRunProperties(){ FontSize = 2000, Kerning = 1200 };

            A.SolidFill solidFill8 = new A.SolidFill();
            A.SchemeColor schemeColor15 = new A.SchemeColor(){ Val = A.SchemeColorValues.Text1 };

            solidFill8.Append(schemeColor15);
            A.LatinFont latinFont3 = new A.LatinFont(){ Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont3 = new A.EastAsianFont(){ Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont3 = new A.ComplexScriptFont(){ Typeface = "+mn-cs" };

            defaultRunProperties25.Append(solidFill8);
            defaultRunProperties25.Append(latinFont3);
            defaultRunProperties25.Append(eastAsianFont3);
            defaultRunProperties25.Append(complexScriptFont3);

            level2ParagraphProperties3.Append(lineSpacing3);
            level2ParagraphProperties3.Append(spaceBefore3);
            level2ParagraphProperties3.Append(bulletFont2);
            level2ParagraphProperties3.Append(characterBullet2);
            level2ParagraphProperties3.Append(defaultRunProperties25);

            A.Level3ParagraphProperties level3ParagraphProperties3 = new A.Level3ParagraphProperties(){ LeftMargin = 952462, Indent = -190492, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 761970, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.LineSpacing lineSpacing4 = new A.LineSpacing();
            A.SpacingPercent spacingPercent5 = new A.SpacingPercent(){ Val = 90000 };

            lineSpacing4.Append(spacingPercent5);

            A.SpaceBefore spaceBefore4 = new A.SpaceBefore();
            A.SpacingPoints spacingPoints3 = new A.SpacingPoints(){ Val = 417 };

            spaceBefore4.Append(spacingPoints3);
            A.BulletFont bulletFont3 = new A.BulletFont(){ Typeface = "Arial", Panose = "020B0604020202020204", PitchFamily = 34, CharacterSet = 0 };
            A.CharacterBullet characterBullet3 = new A.CharacterBullet(){ Char = "•" };

            A.DefaultRunProperties defaultRunProperties26 = new A.DefaultRunProperties(){ FontSize = 1667, Kerning = 1200 };

            A.SolidFill solidFill9 = new A.SolidFill();
            A.SchemeColor schemeColor16 = new A.SchemeColor(){ Val = A.SchemeColorValues.Text1 };

            solidFill9.Append(schemeColor16);
            A.LatinFont latinFont4 = new A.LatinFont(){ Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont4 = new A.EastAsianFont(){ Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont4 = new A.ComplexScriptFont(){ Typeface = "+mn-cs" };

            defaultRunProperties26.Append(solidFill9);
            defaultRunProperties26.Append(latinFont4);
            defaultRunProperties26.Append(eastAsianFont4);
            defaultRunProperties26.Append(complexScriptFont4);

            level3ParagraphProperties3.Append(lineSpacing4);
            level3ParagraphProperties3.Append(spaceBefore4);
            level3ParagraphProperties3.Append(bulletFont3);
            level3ParagraphProperties3.Append(characterBullet3);
            level3ParagraphProperties3.Append(defaultRunProperties26);

            A.Level4ParagraphProperties level4ParagraphProperties3 = new A.Level4ParagraphProperties(){ LeftMargin = 1333447, Indent = -190492, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 761970, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.LineSpacing lineSpacing5 = new A.LineSpacing();
            A.SpacingPercent spacingPercent6 = new A.SpacingPercent(){ Val = 90000 };

            lineSpacing5.Append(spacingPercent6);

            A.SpaceBefore spaceBefore5 = new A.SpaceBefore();
            A.SpacingPoints spacingPoints4 = new A.SpacingPoints(){ Val = 417 };

            spaceBefore5.Append(spacingPoints4);
            A.BulletFont bulletFont4 = new A.BulletFont(){ Typeface = "Arial", Panose = "020B0604020202020204", PitchFamily = 34, CharacterSet = 0 };
            A.CharacterBullet characterBullet4 = new A.CharacterBullet(){ Char = "•" };

            A.DefaultRunProperties defaultRunProperties27 = new A.DefaultRunProperties(){ FontSize = 1500, Kerning = 1200 };

            A.SolidFill solidFill10 = new A.SolidFill();
            A.SchemeColor schemeColor17 = new A.SchemeColor(){ Val = A.SchemeColorValues.Text1 };

            solidFill10.Append(schemeColor17);
            A.LatinFont latinFont5 = new A.LatinFont(){ Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont5 = new A.EastAsianFont(){ Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont5 = new A.ComplexScriptFont(){ Typeface = "+mn-cs" };

            defaultRunProperties27.Append(solidFill10);
            defaultRunProperties27.Append(latinFont5);
            defaultRunProperties27.Append(eastAsianFont5);
            defaultRunProperties27.Append(complexScriptFont5);

            level4ParagraphProperties3.Append(lineSpacing5);
            level4ParagraphProperties3.Append(spaceBefore5);
            level4ParagraphProperties3.Append(bulletFont4);
            level4ParagraphProperties3.Append(characterBullet4);
            level4ParagraphProperties3.Append(defaultRunProperties27);

            A.Level5ParagraphProperties level5ParagraphProperties3 = new A.Level5ParagraphProperties(){ LeftMargin = 1714431, Indent = -190492, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 761970, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.LineSpacing lineSpacing6 = new A.LineSpacing();
            A.SpacingPercent spacingPercent7 = new A.SpacingPercent(){ Val = 90000 };

            lineSpacing6.Append(spacingPercent7);

            A.SpaceBefore spaceBefore6 = new A.SpaceBefore();
            A.SpacingPoints spacingPoints5 = new A.SpacingPoints(){ Val = 417 };

            spaceBefore6.Append(spacingPoints5);
            A.BulletFont bulletFont5 = new A.BulletFont(){ Typeface = "Arial", Panose = "020B0604020202020204", PitchFamily = 34, CharacterSet = 0 };
            A.CharacterBullet characterBullet5 = new A.CharacterBullet(){ Char = "•" };

            A.DefaultRunProperties defaultRunProperties28 = new A.DefaultRunProperties(){ FontSize = 1500, Kerning = 1200 };

            A.SolidFill solidFill11 = new A.SolidFill();
            A.SchemeColor schemeColor18 = new A.SchemeColor(){ Val = A.SchemeColorValues.Text1 };

            solidFill11.Append(schemeColor18);
            A.LatinFont latinFont6 = new A.LatinFont(){ Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont6 = new A.EastAsianFont(){ Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont6 = new A.ComplexScriptFont(){ Typeface = "+mn-cs" };

            defaultRunProperties28.Append(solidFill11);
            defaultRunProperties28.Append(latinFont6);
            defaultRunProperties28.Append(eastAsianFont6);
            defaultRunProperties28.Append(complexScriptFont6);

            level5ParagraphProperties3.Append(lineSpacing6);
            level5ParagraphProperties3.Append(spaceBefore6);
            level5ParagraphProperties3.Append(bulletFont5);
            level5ParagraphProperties3.Append(characterBullet5);
            level5ParagraphProperties3.Append(defaultRunProperties28);

            A.Level6ParagraphProperties level6ParagraphProperties3 = new A.Level6ParagraphProperties(){ LeftMargin = 2095416, Indent = -190492, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 761970, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.LineSpacing lineSpacing7 = new A.LineSpacing();
            A.SpacingPercent spacingPercent8 = new A.SpacingPercent(){ Val = 90000 };

            lineSpacing7.Append(spacingPercent8);

            A.SpaceBefore spaceBefore7 = new A.SpaceBefore();
            A.SpacingPoints spacingPoints6 = new A.SpacingPoints(){ Val = 417 };

            spaceBefore7.Append(spacingPoints6);
            A.BulletFont bulletFont6 = new A.BulletFont(){ Typeface = "Arial", Panose = "020B0604020202020204", PitchFamily = 34, CharacterSet = 0 };
            A.CharacterBullet characterBullet6 = new A.CharacterBullet(){ Char = "•" };

            A.DefaultRunProperties defaultRunProperties29 = new A.DefaultRunProperties(){ FontSize = 1500, Kerning = 1200 };

            A.SolidFill solidFill12 = new A.SolidFill();
            A.SchemeColor schemeColor19 = new A.SchemeColor(){ Val = A.SchemeColorValues.Text1 };

            solidFill12.Append(schemeColor19);
            A.LatinFont latinFont7 = new A.LatinFont(){ Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont7 = new A.EastAsianFont(){ Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont7 = new A.ComplexScriptFont(){ Typeface = "+mn-cs" };

            defaultRunProperties29.Append(solidFill12);
            defaultRunProperties29.Append(latinFont7);
            defaultRunProperties29.Append(eastAsianFont7);
            defaultRunProperties29.Append(complexScriptFont7);

            level6ParagraphProperties3.Append(lineSpacing7);
            level6ParagraphProperties3.Append(spaceBefore7);
            level6ParagraphProperties3.Append(bulletFont6);
            level6ParagraphProperties3.Append(characterBullet6);
            level6ParagraphProperties3.Append(defaultRunProperties29);

            A.Level7ParagraphProperties level7ParagraphProperties3 = new A.Level7ParagraphProperties(){ LeftMargin = 2476401, Indent = -190492, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 761970, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.LineSpacing lineSpacing8 = new A.LineSpacing();
            A.SpacingPercent spacingPercent9 = new A.SpacingPercent(){ Val = 90000 };

            lineSpacing8.Append(spacingPercent9);

            A.SpaceBefore spaceBefore8 = new A.SpaceBefore();
            A.SpacingPoints spacingPoints7 = new A.SpacingPoints(){ Val = 417 };

            spaceBefore8.Append(spacingPoints7);
            A.BulletFont bulletFont7 = new A.BulletFont(){ Typeface = "Arial", Panose = "020B0604020202020204", PitchFamily = 34, CharacterSet = 0 };
            A.CharacterBullet characterBullet7 = new A.CharacterBullet(){ Char = "•" };

            A.DefaultRunProperties defaultRunProperties30 = new A.DefaultRunProperties(){ FontSize = 1500, Kerning = 1200 };

            A.SolidFill solidFill13 = new A.SolidFill();
            A.SchemeColor schemeColor20 = new A.SchemeColor(){ Val = A.SchemeColorValues.Text1 };

            solidFill13.Append(schemeColor20);
            A.LatinFont latinFont8 = new A.LatinFont(){ Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont8 = new A.EastAsianFont(){ Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont8 = new A.ComplexScriptFont(){ Typeface = "+mn-cs" };

            defaultRunProperties30.Append(solidFill13);
            defaultRunProperties30.Append(latinFont8);
            defaultRunProperties30.Append(eastAsianFont8);
            defaultRunProperties30.Append(complexScriptFont8);

            level7ParagraphProperties3.Append(lineSpacing8);
            level7ParagraphProperties3.Append(spaceBefore8);
            level7ParagraphProperties3.Append(bulletFont7);
            level7ParagraphProperties3.Append(characterBullet7);
            level7ParagraphProperties3.Append(defaultRunProperties30);

            A.Level8ParagraphProperties level8ParagraphProperties3 = new A.Level8ParagraphProperties(){ LeftMargin = 2857386, Indent = -190492, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 761970, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.LineSpacing lineSpacing9 = new A.LineSpacing();
            A.SpacingPercent spacingPercent10 = new A.SpacingPercent(){ Val = 90000 };

            lineSpacing9.Append(spacingPercent10);

            A.SpaceBefore spaceBefore9 = new A.SpaceBefore();
            A.SpacingPoints spacingPoints8 = new A.SpacingPoints(){ Val = 417 };

            spaceBefore9.Append(spacingPoints8);
            A.BulletFont bulletFont8 = new A.BulletFont(){ Typeface = "Arial", Panose = "020B0604020202020204", PitchFamily = 34, CharacterSet = 0 };
            A.CharacterBullet characterBullet8 = new A.CharacterBullet(){ Char = "•" };

            A.DefaultRunProperties defaultRunProperties31 = new A.DefaultRunProperties(){ FontSize = 1500, Kerning = 1200 };

            A.SolidFill solidFill14 = new A.SolidFill();
            A.SchemeColor schemeColor21 = new A.SchemeColor(){ Val = A.SchemeColorValues.Text1 };

            solidFill14.Append(schemeColor21);
            A.LatinFont latinFont9 = new A.LatinFont(){ Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont9 = new A.EastAsianFont(){ Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont9 = new A.ComplexScriptFont(){ Typeface = "+mn-cs" };

            defaultRunProperties31.Append(solidFill14);
            defaultRunProperties31.Append(latinFont9);
            defaultRunProperties31.Append(eastAsianFont9);
            defaultRunProperties31.Append(complexScriptFont9);

            level8ParagraphProperties3.Append(lineSpacing9);
            level8ParagraphProperties3.Append(spaceBefore9);
            level8ParagraphProperties3.Append(bulletFont8);
            level8ParagraphProperties3.Append(characterBullet8);
            level8ParagraphProperties3.Append(defaultRunProperties31);

            A.Level9ParagraphProperties level9ParagraphProperties3 = new A.Level9ParagraphProperties(){ LeftMargin = 3238370, Indent = -190492, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 761970, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.LineSpacing lineSpacing10 = new A.LineSpacing();
            A.SpacingPercent spacingPercent11 = new A.SpacingPercent(){ Val = 90000 };

            lineSpacing10.Append(spacingPercent11);

            A.SpaceBefore spaceBefore10 = new A.SpaceBefore();
            A.SpacingPoints spacingPoints9 = new A.SpacingPoints(){ Val = 417 };

            spaceBefore10.Append(spacingPoints9);
            A.BulletFont bulletFont9 = new A.BulletFont(){ Typeface = "Arial", Panose = "020B0604020202020204", PitchFamily = 34, CharacterSet = 0 };
            A.CharacterBullet characterBullet9 = new A.CharacterBullet(){ Char = "•" };

            A.DefaultRunProperties defaultRunProperties32 = new A.DefaultRunProperties(){ FontSize = 1500, Kerning = 1200 };

            A.SolidFill solidFill15 = new A.SolidFill();
            A.SchemeColor schemeColor22 = new A.SchemeColor(){ Val = A.SchemeColorValues.Text1 };

            solidFill15.Append(schemeColor22);
            A.LatinFont latinFont10 = new A.LatinFont(){ Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont10 = new A.EastAsianFont(){ Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont10 = new A.ComplexScriptFont(){ Typeface = "+mn-cs" };

            defaultRunProperties32.Append(solidFill15);
            defaultRunProperties32.Append(latinFont10);
            defaultRunProperties32.Append(eastAsianFont10);
            defaultRunProperties32.Append(complexScriptFont10);

            level9ParagraphProperties3.Append(lineSpacing10);
            level9ParagraphProperties3.Append(spaceBefore10);
            level9ParagraphProperties3.Append(bulletFont9);
            level9ParagraphProperties3.Append(characterBullet9);
            level9ParagraphProperties3.Append(defaultRunProperties32);

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

            A.DefaultParagraphProperties defaultParagraphProperties1 = new A.DefaultParagraphProperties();
            A.DefaultRunProperties defaultRunProperties33 = new A.DefaultRunProperties(){ Language = "en-US" };

            defaultParagraphProperties1.Append(defaultRunProperties33);

            A.Level1ParagraphProperties level1ParagraphProperties9 = new A.Level1ParagraphProperties(){ LeftMargin = 0, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 761970, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.DefaultRunProperties defaultRunProperties34 = new A.DefaultRunProperties(){ FontSize = 1500, Kerning = 1200 };

            A.SolidFill solidFill16 = new A.SolidFill();
            A.SchemeColor schemeColor23 = new A.SchemeColor(){ Val = A.SchemeColorValues.Text1 };

            solidFill16.Append(schemeColor23);
            A.LatinFont latinFont11 = new A.LatinFont(){ Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont11 = new A.EastAsianFont(){ Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont11 = new A.ComplexScriptFont(){ Typeface = "+mn-cs" };

            defaultRunProperties34.Append(solidFill16);
            defaultRunProperties34.Append(latinFont11);
            defaultRunProperties34.Append(eastAsianFont11);
            defaultRunProperties34.Append(complexScriptFont11);

            level1ParagraphProperties9.Append(defaultRunProperties34);

            A.Level2ParagraphProperties level2ParagraphProperties4 = new A.Level2ParagraphProperties(){ LeftMargin = 380985, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 761970, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.DefaultRunProperties defaultRunProperties35 = new A.DefaultRunProperties(){ FontSize = 1500, Kerning = 1200 };

            A.SolidFill solidFill17 = new A.SolidFill();
            A.SchemeColor schemeColor24 = new A.SchemeColor(){ Val = A.SchemeColorValues.Text1 };

            solidFill17.Append(schemeColor24);
            A.LatinFont latinFont12 = new A.LatinFont(){ Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont12 = new A.EastAsianFont(){ Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont12 = new A.ComplexScriptFont(){ Typeface = "+mn-cs" };

            defaultRunProperties35.Append(solidFill17);
            defaultRunProperties35.Append(latinFont12);
            defaultRunProperties35.Append(eastAsianFont12);
            defaultRunProperties35.Append(complexScriptFont12);

            level2ParagraphProperties4.Append(defaultRunProperties35);

            A.Level3ParagraphProperties level3ParagraphProperties4 = new A.Level3ParagraphProperties(){ LeftMargin = 761970, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 761970, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.DefaultRunProperties defaultRunProperties36 = new A.DefaultRunProperties(){ FontSize = 1500, Kerning = 1200 };

            A.SolidFill solidFill18 = new A.SolidFill();
            A.SchemeColor schemeColor25 = new A.SchemeColor(){ Val = A.SchemeColorValues.Text1 };

            solidFill18.Append(schemeColor25);
            A.LatinFont latinFont13 = new A.LatinFont(){ Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont13 = new A.EastAsianFont(){ Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont13 = new A.ComplexScriptFont(){ Typeface = "+mn-cs" };

            defaultRunProperties36.Append(solidFill18);
            defaultRunProperties36.Append(latinFont13);
            defaultRunProperties36.Append(eastAsianFont13);
            defaultRunProperties36.Append(complexScriptFont13);

            level3ParagraphProperties4.Append(defaultRunProperties36);

            A.Level4ParagraphProperties level4ParagraphProperties4 = new A.Level4ParagraphProperties(){ LeftMargin = 1142954, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 761970, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.DefaultRunProperties defaultRunProperties37 = new A.DefaultRunProperties(){ FontSize = 1500, Kerning = 1200 };

            A.SolidFill solidFill19 = new A.SolidFill();
            A.SchemeColor schemeColor26 = new A.SchemeColor(){ Val = A.SchemeColorValues.Text1 };

            solidFill19.Append(schemeColor26);
            A.LatinFont latinFont14 = new A.LatinFont(){ Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont14 = new A.EastAsianFont(){ Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont14 = new A.ComplexScriptFont(){ Typeface = "+mn-cs" };

            defaultRunProperties37.Append(solidFill19);
            defaultRunProperties37.Append(latinFont14);
            defaultRunProperties37.Append(eastAsianFont14);
            defaultRunProperties37.Append(complexScriptFont14);

            level4ParagraphProperties4.Append(defaultRunProperties37);

            A.Level5ParagraphProperties level5ParagraphProperties4 = new A.Level5ParagraphProperties(){ LeftMargin = 1523939, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 761970, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.DefaultRunProperties defaultRunProperties38 = new A.DefaultRunProperties(){ FontSize = 1500, Kerning = 1200 };

            A.SolidFill solidFill20 = new A.SolidFill();
            A.SchemeColor schemeColor27 = new A.SchemeColor(){ Val = A.SchemeColorValues.Text1 };

            solidFill20.Append(schemeColor27);
            A.LatinFont latinFont15 = new A.LatinFont(){ Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont15 = new A.EastAsianFont(){ Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont15 = new A.ComplexScriptFont(){ Typeface = "+mn-cs" };

            defaultRunProperties38.Append(solidFill20);
            defaultRunProperties38.Append(latinFont15);
            defaultRunProperties38.Append(eastAsianFont15);
            defaultRunProperties38.Append(complexScriptFont15);

            level5ParagraphProperties4.Append(defaultRunProperties38);

            A.Level6ParagraphProperties level6ParagraphProperties4 = new A.Level6ParagraphProperties(){ LeftMargin = 1904924, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 761970, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.DefaultRunProperties defaultRunProperties39 = new A.DefaultRunProperties(){ FontSize = 1500, Kerning = 1200 };

            A.SolidFill solidFill21 = new A.SolidFill();
            A.SchemeColor schemeColor28 = new A.SchemeColor(){ Val = A.SchemeColorValues.Text1 };

            solidFill21.Append(schemeColor28);
            A.LatinFont latinFont16 = new A.LatinFont(){ Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont16 = new A.EastAsianFont(){ Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont16 = new A.ComplexScriptFont(){ Typeface = "+mn-cs" };

            defaultRunProperties39.Append(solidFill21);
            defaultRunProperties39.Append(latinFont16);
            defaultRunProperties39.Append(eastAsianFont16);
            defaultRunProperties39.Append(complexScriptFont16);

            level6ParagraphProperties4.Append(defaultRunProperties39);

            A.Level7ParagraphProperties level7ParagraphProperties4 = new A.Level7ParagraphProperties(){ LeftMargin = 2285909, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 761970, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.DefaultRunProperties defaultRunProperties40 = new A.DefaultRunProperties(){ FontSize = 1500, Kerning = 1200 };

            A.SolidFill solidFill22 = new A.SolidFill();
            A.SchemeColor schemeColor29 = new A.SchemeColor(){ Val = A.SchemeColorValues.Text1 };

            solidFill22.Append(schemeColor29);
            A.LatinFont latinFont17 = new A.LatinFont(){ Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont17 = new A.EastAsianFont(){ Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont17 = new A.ComplexScriptFont(){ Typeface = "+mn-cs" };

            defaultRunProperties40.Append(solidFill22);
            defaultRunProperties40.Append(latinFont17);
            defaultRunProperties40.Append(eastAsianFont17);
            defaultRunProperties40.Append(complexScriptFont17);

            level7ParagraphProperties4.Append(defaultRunProperties40);

            A.Level8ParagraphProperties level8ParagraphProperties4 = new A.Level8ParagraphProperties(){ LeftMargin = 2666893, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 761970, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.DefaultRunProperties defaultRunProperties41 = new A.DefaultRunProperties(){ FontSize = 1500, Kerning = 1200 };

            A.SolidFill solidFill23 = new A.SolidFill();
            A.SchemeColor schemeColor30 = new A.SchemeColor(){ Val = A.SchemeColorValues.Text1 };

            solidFill23.Append(schemeColor30);
            A.LatinFont latinFont18 = new A.LatinFont(){ Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont18 = new A.EastAsianFont(){ Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont18 = new A.ComplexScriptFont(){ Typeface = "+mn-cs" };

            defaultRunProperties41.Append(solidFill23);
            defaultRunProperties41.Append(latinFont18);
            defaultRunProperties41.Append(eastAsianFont18);
            defaultRunProperties41.Append(complexScriptFont18);

            level8ParagraphProperties4.Append(defaultRunProperties41);

            A.Level9ParagraphProperties level9ParagraphProperties4 = new A.Level9ParagraphProperties(){ LeftMargin = 3047878, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 761970, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.DefaultRunProperties defaultRunProperties42 = new A.DefaultRunProperties(){ FontSize = 1500, Kerning = 1200 };

            A.SolidFill solidFill24 = new A.SolidFill();
            A.SchemeColor schemeColor31 = new A.SchemeColor(){ Val = A.SchemeColorValues.Text1 };

            solidFill24.Append(schemeColor31);
            A.LatinFont latinFont19 = new A.LatinFont(){ Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont19 = new A.EastAsianFont(){ Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont19 = new A.ComplexScriptFont(){ Typeface = "+mn-cs" };

            defaultRunProperties42.Append(solidFill24);
            defaultRunProperties42.Append(latinFont19);
            defaultRunProperties42.Append(eastAsianFont19);
            defaultRunProperties42.Append(complexScriptFont19);

            level9ParagraphProperties4.Append(defaultRunProperties42);

            otherStyle1.Append(defaultParagraphProperties1);
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

            SlideMasterExtensionList slideMasterExtensionList1 = new SlideMasterExtensionList();

            SlideMasterExtension slideMasterExtension1 = new SlideMasterExtension(){ Uri = "{27BBF7A9-308A-43DC-89C8-2F10F3537804}" };

            P15.SlideGuideList slideGuideList1 = new P15.SlideGuideList();
            slideGuideList1.AddNamespaceDeclaration("p15", "http://schemas.microsoft.com/office/powerpoint/2012/main");

            P15.ExtendedGuide extendedGuide1 = new P15.ExtendedGuide(){ Id = (UInt32Value)1U, Orientation = DirectionValues.Horizontal, Position = 99 };

            P15.ColorType colorType1 = new P15.ColorType();
            A.RgbColorModelHex rgbColorModelHex1 = new A.RgbColorModelHex(){ Val = "F26B43" };

            colorType1.Append(rgbColorModelHex1);

            extendedGuide1.Append(colorType1);

            P15.ExtendedGuide extendedGuide2 = new P15.ExtendedGuide(){ Id = (UInt32Value)2U, Position = 170 };

            P15.ColorType colorType2 = new P15.ColorType();
            A.RgbColorModelHex rgbColorModelHex2 = new A.RgbColorModelHex(){ Val = "F26B43" };

            colorType2.Append(rgbColorModelHex2);

            extendedGuide2.Append(colorType2);

            P15.ExtendedGuide extendedGuide3 = new P15.ExtendedGuide(){ Id = (UInt32Value)3U, Position = 4649 };

            P15.ColorType colorType3 = new P15.ColorType();
            A.RgbColorModelHex rgbColorModelHex3 = new A.RgbColorModelHex(){ Val = "F26B43" };

            colorType3.Append(rgbColorModelHex3);

            extendedGuide3.Append(colorType3);

            P15.ExtendedGuide extendedGuide4 = new P15.ExtendedGuide(){ Id = (UInt32Value)4U, Orientation = DirectionValues.Horizontal, Position = 689 };

            P15.ColorType colorType4 = new P15.ColorType();
            A.RgbColorModelHex rgbColorModelHex4 = new A.RgbColorModelHex(){ Val = "F26B43" };

            colorType4.Append(rgbColorModelHex4);

            extendedGuide4.Append(colorType4);

            P15.ExtendedGuide extendedGuide5 = new P15.ExtendedGuide(){ Id = (UInt32Value)5U, Orientation = DirectionValues.Horizontal, Position = 839 };

            P15.ColorType colorType5 = new P15.ColorType();
            A.RgbColorModelHex rgbColorModelHex5 = new A.RgbColorModelHex(){ Val = "F26B43" };

            colorType5.Append(rgbColorModelHex5);

            extendedGuide5.Append(colorType5);

            P15.ExtendedGuide extendedGuide6 = new P15.ExtendedGuide(){ Id = (UInt32Value)6U, Orientation = DirectionValues.Horizontal, Position = 3161 };

            P15.ColorType colorType6 = new P15.ColorType();
            A.RgbColorModelHex rgbColorModelHex6 = new A.RgbColorModelHex(){ Val = "F26B43" };

            colorType6.Append(rgbColorModelHex6);

            extendedGuide6.Append(colorType6);

            P15.ExtendedGuide extendedGuide7 = new P15.ExtendedGuide(){ Id = (UInt32Value)7U, Orientation = DirectionValues.Horizontal, Position = 3342 };

            P15.ColorType colorType7 = new P15.ColorType();
            A.RgbColorModelHex rgbColorModelHex7 = new A.RgbColorModelHex(){ Val = "F26B43" };

            colorType7.Append(rgbColorModelHex7);

            extendedGuide7.Append(colorType7);

            P15.ExtendedGuide extendedGuide8 = new P15.ExtendedGuide(){ Id = (UInt32Value)8U, Orientation = DirectionValues.Horizontal, Position = 3546 };

            P15.ColorType colorType8 = new P15.ColorType();
            A.RgbColorModelHex rgbColorModelHex8 = new A.RgbColorModelHex(){ Val = "F26B43" };

            colorType8.Append(rgbColorModelHex8);

            extendedGuide8.Append(colorType8);

            slideGuideList1.Append(extendedGuide1);
            slideGuideList1.Append(extendedGuide2);
            slideGuideList1.Append(extendedGuide3);
            slideGuideList1.Append(extendedGuide4);
            slideGuideList1.Append(extendedGuide5);
            slideGuideList1.Append(extendedGuide6);
            slideGuideList1.Append(extendedGuide7);
            slideGuideList1.Append(extendedGuide8);

            slideMasterExtension1.Append(slideGuideList1);

            slideMasterExtensionList1.Append(slideMasterExtension1);

            slideMaster1.Append(commonSlideData2);
            slideMaster1.Append(colorMap1);
            slideMaster1.Append(slideLayoutIdList1);
            slideMaster1.Append(headerFooter1);
            slideMaster1.Append(textStyles1);
            slideMaster1.Append(slideMasterExtensionList1);

            slideMasterPart1.SlideMaster = slideMaster1;
        }

        // Generates content of themePart1.
        private void GenerateThemePart1Content(ThemePart themePart1)
        {
            A.Theme theme1 = new A.Theme(){ Name = "Office Theme" };
            theme1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

            A.ThemeElements themeElements1 = new A.ThemeElements();

            A.ColorScheme colorScheme1 = new A.ColorScheme(){ Name = "Office Theme" };

            A.Dark1Color dark1Color1 = new A.Dark1Color();
            A.SystemColor systemColor1 = new A.SystemColor(){ Val = A.SystemColorValues.WindowText, LastColor = "000000" };

            dark1Color1.Append(systemColor1);

            A.Light1Color light1Color1 = new A.Light1Color();
            A.SystemColor systemColor2 = new A.SystemColor(){ Val = A.SystemColorValues.Window, LastColor = "FFFFFF" };

            light1Color1.Append(systemColor2);

            A.Dark2Color dark2Color1 = new A.Dark2Color();
            A.RgbColorModelHex rgbColorModelHex9 = new A.RgbColorModelHex(){ Val = "44546A" };

            dark2Color1.Append(rgbColorModelHex9);

            A.Light2Color light2Color1 = new A.Light2Color();
            A.RgbColorModelHex rgbColorModelHex10 = new A.RgbColorModelHex(){ Val = "E7E6E6" };

            light2Color1.Append(rgbColorModelHex10);

            A.Accent1Color accent1Color1 = new A.Accent1Color();
            A.RgbColorModelHex rgbColorModelHex11 = new A.RgbColorModelHex(){ Val = "4472C4" };

            accent1Color1.Append(rgbColorModelHex11);

            A.Accent2Color accent2Color1 = new A.Accent2Color();
            A.RgbColorModelHex rgbColorModelHex12 = new A.RgbColorModelHex(){ Val = "ED7D31" };

            accent2Color1.Append(rgbColorModelHex12);

            A.Accent3Color accent3Color1 = new A.Accent3Color();
            A.RgbColorModelHex rgbColorModelHex13 = new A.RgbColorModelHex(){ Val = "A5A5A5" };

            accent3Color1.Append(rgbColorModelHex13);

            A.Accent4Color accent4Color1 = new A.Accent4Color();
            A.RgbColorModelHex rgbColorModelHex14 = new A.RgbColorModelHex(){ Val = "FFC000" };

            accent4Color1.Append(rgbColorModelHex14);

            A.Accent5Color accent5Color1 = new A.Accent5Color();
            A.RgbColorModelHex rgbColorModelHex15 = new A.RgbColorModelHex(){ Val = "5B9BD5" };

            accent5Color1.Append(rgbColorModelHex15);

            A.Accent6Color accent6Color1 = new A.Accent6Color();
            A.RgbColorModelHex rgbColorModelHex16 = new A.RgbColorModelHex(){ Val = "70AD47" };

            accent6Color1.Append(rgbColorModelHex16);

            A.Hyperlink hyperlink1 = new A.Hyperlink();
            A.RgbColorModelHex rgbColorModelHex17 = new A.RgbColorModelHex(){ Val = "0563C1" };

            hyperlink1.Append(rgbColorModelHex17);

            A.FollowedHyperlinkColor followedHyperlinkColor1 = new A.FollowedHyperlinkColor();
            A.RgbColorModelHex rgbColorModelHex18 = new A.RgbColorModelHex(){ Val = "954F72" };

            followedHyperlinkColor1.Append(rgbColorModelHex18);

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

            A.FontScheme fontScheme1 = new A.FontScheme(){ Name = "Office Theme" };

            A.MajorFont majorFont1 = new A.MajorFont();
            A.LatinFont latinFont20 = new A.LatinFont(){ Typeface = "Calibri Light", Panose = "020F0302020204030204" };
            A.EastAsianFont eastAsianFont20 = new A.EastAsianFont(){ Typeface = "" };
            A.ComplexScriptFont complexScriptFont20 = new A.ComplexScriptFont(){ Typeface = "" };
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

            majorFont1.Append(latinFont20);
            majorFont1.Append(eastAsianFont20);
            majorFont1.Append(complexScriptFont20);
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
            A.LatinFont latinFont21 = new A.LatinFont(){ Typeface = "Calibri", Panose = "020F0502020204030204" };
            A.EastAsianFont eastAsianFont21 = new A.EastAsianFont(){ Typeface = "" };
            A.ComplexScriptFont complexScriptFont21 = new A.ComplexScriptFont(){ Typeface = "" };
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

            minorFont1.Append(latinFont21);
            minorFont1.Append(eastAsianFont21);
            minorFont1.Append(complexScriptFont21);
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

            A.FormatScheme formatScheme1 = new A.FormatScheme(){ Name = "Office Theme" };

            A.FillStyleList fillStyleList1 = new A.FillStyleList();

            A.SolidFill solidFill25 = new A.SolidFill();
            A.SchemeColor schemeColor32 = new A.SchemeColor(){ Val = A.SchemeColorValues.PhColor };

            solidFill25.Append(schemeColor32);

            A.GradientFill gradientFill2 = new A.GradientFill(){ RotateWithShape = true };

            A.GradientStopList gradientStopList2 = new A.GradientStopList();

            A.GradientStop gradientStop3 = new A.GradientStop(){ Position = 0 };

            A.SchemeColor schemeColor33 = new A.SchemeColor(){ Val = A.SchemeColorValues.PhColor };
            A.LuminanceModulation luminanceModulation1 = new A.LuminanceModulation(){ Val = 110000 };
            A.SaturationModulation saturationModulation1 = new A.SaturationModulation(){ Val = 105000 };
            A.Tint tint4 = new A.Tint(){ Val = 67000 };

            schemeColor33.Append(luminanceModulation1);
            schemeColor33.Append(saturationModulation1);
            schemeColor33.Append(tint4);

            gradientStop3.Append(schemeColor33);

            A.GradientStop gradientStop4 = new A.GradientStop(){ Position = 50000 };

            A.SchemeColor schemeColor34 = new A.SchemeColor(){ Val = A.SchemeColorValues.PhColor };
            A.LuminanceModulation luminanceModulation2 = new A.LuminanceModulation(){ Val = 105000 };
            A.SaturationModulation saturationModulation2 = new A.SaturationModulation(){ Val = 103000 };
            A.Tint tint5 = new A.Tint(){ Val = 73000 };

            schemeColor34.Append(luminanceModulation2);
            schemeColor34.Append(saturationModulation2);
            schemeColor34.Append(tint5);

            gradientStop4.Append(schemeColor34);

            A.GradientStop gradientStop5 = new A.GradientStop(){ Position = 100000 };

            A.SchemeColor schemeColor35 = new A.SchemeColor(){ Val = A.SchemeColorValues.PhColor };
            A.LuminanceModulation luminanceModulation3 = new A.LuminanceModulation(){ Val = 105000 };
            A.SaturationModulation saturationModulation3 = new A.SaturationModulation(){ Val = 109000 };
            A.Tint tint6 = new A.Tint(){ Val = 81000 };

            schemeColor35.Append(luminanceModulation3);
            schemeColor35.Append(saturationModulation3);
            schemeColor35.Append(tint6);

            gradientStop5.Append(schemeColor35);

            gradientStopList2.Append(gradientStop3);
            gradientStopList2.Append(gradientStop4);
            gradientStopList2.Append(gradientStop5);
            A.LinearGradientFill linearGradientFill2 = new A.LinearGradientFill(){ Angle = 5400000, Scaled = false };

            gradientFill2.Append(gradientStopList2);
            gradientFill2.Append(linearGradientFill2);

            A.GradientFill gradientFill3 = new A.GradientFill(){ RotateWithShape = true };

            A.GradientStopList gradientStopList3 = new A.GradientStopList();

            A.GradientStop gradientStop6 = new A.GradientStop(){ Position = 0 };

            A.SchemeColor schemeColor36 = new A.SchemeColor(){ Val = A.SchemeColorValues.PhColor };
            A.SaturationModulation saturationModulation4 = new A.SaturationModulation(){ Val = 103000 };
            A.LuminanceModulation luminanceModulation4 = new A.LuminanceModulation(){ Val = 102000 };
            A.Tint tint7 = new A.Tint(){ Val = 94000 };

            schemeColor36.Append(saturationModulation4);
            schemeColor36.Append(luminanceModulation4);
            schemeColor36.Append(tint7);

            gradientStop6.Append(schemeColor36);

            A.GradientStop gradientStop7 = new A.GradientStop(){ Position = 50000 };

            A.SchemeColor schemeColor37 = new A.SchemeColor(){ Val = A.SchemeColorValues.PhColor };
            A.SaturationModulation saturationModulation5 = new A.SaturationModulation(){ Val = 110000 };
            A.LuminanceModulation luminanceModulation5 = new A.LuminanceModulation(){ Val = 100000 };
            A.Shade shade2 = new A.Shade(){ Val = 100000 };

            schemeColor37.Append(saturationModulation5);
            schemeColor37.Append(luminanceModulation5);
            schemeColor37.Append(shade2);

            gradientStop7.Append(schemeColor37);

            A.GradientStop gradientStop8 = new A.GradientStop(){ Position = 100000 };

            A.SchemeColor schemeColor38 = new A.SchemeColor(){ Val = A.SchemeColorValues.PhColor };
            A.LuminanceModulation luminanceModulation6 = new A.LuminanceModulation(){ Val = 99000 };
            A.SaturationModulation saturationModulation6 = new A.SaturationModulation(){ Val = 120000 };
            A.Shade shade3 = new A.Shade(){ Val = 78000 };

            schemeColor38.Append(luminanceModulation6);
            schemeColor38.Append(saturationModulation6);
            schemeColor38.Append(shade3);

            gradientStop8.Append(schemeColor38);

            gradientStopList3.Append(gradientStop6);
            gradientStopList3.Append(gradientStop7);
            gradientStopList3.Append(gradientStop8);
            A.LinearGradientFill linearGradientFill3 = new A.LinearGradientFill(){ Angle = 5400000, Scaled = false };

            gradientFill3.Append(gradientStopList3);
            gradientFill3.Append(linearGradientFill3);

            fillStyleList1.Append(solidFill25);
            fillStyleList1.Append(gradientFill2);
            fillStyleList1.Append(gradientFill3);

            A.LineStyleList lineStyleList1 = new A.LineStyleList();

            A.Outline outline2 = new A.Outline(){ Width = 6350, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill26 = new A.SolidFill();
            A.SchemeColor schemeColor39 = new A.SchemeColor(){ Val = A.SchemeColorValues.PhColor };

            solidFill26.Append(schemeColor39);
            A.PresetDash presetDash1 = new A.PresetDash(){ Val = A.PresetLineDashValues.Solid };
            A.Miter miter1 = new A.Miter(){ Limit = 800000 };

            outline2.Append(solidFill26);
            outline2.Append(presetDash1);
            outline2.Append(miter1);

            A.Outline outline3 = new A.Outline(){ Width = 12700, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill27 = new A.SolidFill();
            A.SchemeColor schemeColor40 = new A.SchemeColor(){ Val = A.SchemeColorValues.PhColor };

            solidFill27.Append(schemeColor40);
            A.PresetDash presetDash2 = new A.PresetDash(){ Val = A.PresetLineDashValues.Solid };
            A.Miter miter2 = new A.Miter(){ Limit = 800000 };

            outline3.Append(solidFill27);
            outline3.Append(presetDash2);
            outline3.Append(miter2);

            A.Outline outline4 = new A.Outline(){ Width = 19050, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill28 = new A.SolidFill();
            A.SchemeColor schemeColor41 = new A.SchemeColor(){ Val = A.SchemeColorValues.PhColor };

            solidFill28.Append(schemeColor41);
            A.PresetDash presetDash3 = new A.PresetDash(){ Val = A.PresetLineDashValues.Solid };
            A.Miter miter3 = new A.Miter(){ Limit = 800000 };

            outline4.Append(solidFill28);
            outline4.Append(presetDash3);
            outline4.Append(miter3);

            lineStyleList1.Append(outline2);
            lineStyleList1.Append(outline3);
            lineStyleList1.Append(outline4);

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

            A.RgbColorModelHex rgbColorModelHex19 = new A.RgbColorModelHex(){ Val = "000000" };
            A.Alpha alpha1 = new A.Alpha(){ Val = 63000 };

            rgbColorModelHex19.Append(alpha1);

            outerShadow1.Append(rgbColorModelHex19);

            effectList3.Append(outerShadow1);

            effectStyle3.Append(effectList3);

            effectStyleList1.Append(effectStyle1);
            effectStyleList1.Append(effectStyle2);
            effectStyleList1.Append(effectStyle3);

            A.BackgroundFillStyleList backgroundFillStyleList1 = new A.BackgroundFillStyleList();

            A.SolidFill solidFill29 = new A.SolidFill();
            A.SchemeColor schemeColor42 = new A.SchemeColor(){ Val = A.SchemeColorValues.PhColor };

            solidFill29.Append(schemeColor42);

            A.SolidFill solidFill30 = new A.SolidFill();

            A.SchemeColor schemeColor43 = new A.SchemeColor(){ Val = A.SchemeColorValues.PhColor };
            A.Tint tint8 = new A.Tint(){ Val = 95000 };
            A.SaturationModulation saturationModulation7 = new A.SaturationModulation(){ Val = 170000 };

            schemeColor43.Append(tint8);
            schemeColor43.Append(saturationModulation7);

            solidFill30.Append(schemeColor43);

            A.GradientFill gradientFill4 = new A.GradientFill(){ RotateWithShape = true };

            A.GradientStopList gradientStopList4 = new A.GradientStopList();

            A.GradientStop gradientStop9 = new A.GradientStop(){ Position = 0 };

            A.SchemeColor schemeColor44 = new A.SchemeColor(){ Val = A.SchemeColorValues.PhColor };
            A.Tint tint9 = new A.Tint(){ Val = 93000 };
            A.SaturationModulation saturationModulation8 = new A.SaturationModulation(){ Val = 150000 };
            A.Shade shade4 = new A.Shade(){ Val = 98000 };
            A.LuminanceModulation luminanceModulation7 = new A.LuminanceModulation(){ Val = 102000 };

            schemeColor44.Append(tint9);
            schemeColor44.Append(saturationModulation8);
            schemeColor44.Append(shade4);
            schemeColor44.Append(luminanceModulation7);

            gradientStop9.Append(schemeColor44);

            A.GradientStop gradientStop10 = new A.GradientStop(){ Position = 50000 };

            A.SchemeColor schemeColor45 = new A.SchemeColor(){ Val = A.SchemeColorValues.PhColor };
            A.Tint tint10 = new A.Tint(){ Val = 98000 };
            A.SaturationModulation saturationModulation9 = new A.SaturationModulation(){ Val = 130000 };
            A.Shade shade5 = new A.Shade(){ Val = 90000 };
            A.LuminanceModulation luminanceModulation8 = new A.LuminanceModulation(){ Val = 103000 };

            schemeColor45.Append(tint10);
            schemeColor45.Append(saturationModulation9);
            schemeColor45.Append(shade5);
            schemeColor45.Append(luminanceModulation8);

            gradientStop10.Append(schemeColor45);

            A.GradientStop gradientStop11 = new A.GradientStop(){ Position = 100000 };

            A.SchemeColor schemeColor46 = new A.SchemeColor(){ Val = A.SchemeColorValues.PhColor };
            A.Shade shade6 = new A.Shade(){ Val = 63000 };
            A.SaturationModulation saturationModulation10 = new A.SaturationModulation(){ Val = 120000 };

            schemeColor46.Append(shade6);
            schemeColor46.Append(saturationModulation10);

            gradientStop11.Append(schemeColor46);

            gradientStopList4.Append(gradientStop9);
            gradientStopList4.Append(gradientStop10);
            gradientStopList4.Append(gradientStop11);
            A.LinearGradientFill linearGradientFill4 = new A.LinearGradientFill(){ Angle = 5400000, Scaled = false };

            gradientFill4.Append(gradientStopList4);
            gradientFill4.Append(linearGradientFill4);

            backgroundFillStyleList1.Append(solidFill29);
            backgroundFillStyleList1.Append(solidFill30);
            backgroundFillStyleList1.Append(gradientFill4);

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

            Thm15.ThemeFamily themeFamily1 = new Thm15.ThemeFamily(){ Name = "Office Theme", Id = "{62F939B6-93AF-4DB8-9C6B-D6C7DFDC589F}", Vid = "{4A3C46E8-61CC-4603-A589-7422A47A8E4A}" };
            themeFamily1.AddNamespaceDeclaration("thm15", "http://schemas.microsoft.com/office/thememl/2012/main");

            officeStyleSheetExtension1.Append(themeFamily1);

            officeStyleSheetExtensionList1.Append(officeStyleSheetExtension1);

            theme1.Append(themeElements1);
            theme1.Append(objectDefaults1);
            theme1.Append(extraColorSchemeList1);
            theme1.Append(officeStyleSheetExtensionList1);

            themePart1.Theme = theme1;
        }

        // Generates content of slideLayoutPart2.
        private void GenerateSlideLayoutPart2Content(SlideLayoutPart slideLayoutPart2)
        {
            SlideLayout slideLayout2 = new SlideLayout(){ Type = SlideLayoutValues.SectionHeader, Preserve = true };
            slideLayout2.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            slideLayout2.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            slideLayout2.AddNamespaceDeclaration("p", "http://schemas.openxmlformats.org/presentationml/2006/main");

            CommonSlideData commonSlideData3 = new CommonSlideData(){ Name = "Section Header" };

            ShapeTree shapeTree3 = new ShapeTree();

            NonVisualGroupShapeProperties nonVisualGroupShapeProperties3 = new NonVisualGroupShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties15 = new NonVisualDrawingProperties(){ Id = (UInt32Value)1U, Name = "" };
            NonVisualGroupShapeDrawingProperties nonVisualGroupShapeDrawingProperties3 = new NonVisualGroupShapeDrawingProperties();
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties15 = new ApplicationNonVisualDrawingProperties();

            nonVisualGroupShapeProperties3.Append(nonVisualDrawingProperties15);
            nonVisualGroupShapeProperties3.Append(nonVisualGroupShapeDrawingProperties3);
            nonVisualGroupShapeProperties3.Append(applicationNonVisualDrawingProperties15);

            GroupShapeProperties groupShapeProperties3 = new GroupShapeProperties();

            A.TransformGroup transformGroup3 = new A.TransformGroup();
            A.Offset offset12 = new A.Offset(){ X = 0L, Y = 0L };
            A.Extents extents12 = new A.Extents(){ Cx = 0L, Cy = 0L };
            A.ChildOffset childOffset3 = new A.ChildOffset(){ X = 0L, Y = 0L };
            A.ChildExtents childExtents3 = new A.ChildExtents(){ Cx = 0L, Cy = 0L };

            transformGroup3.Append(offset12);
            transformGroup3.Append(extents12);
            transformGroup3.Append(childOffset3);
            transformGroup3.Append(childExtents3);

            groupShapeProperties3.Append(transformGroup3);

            Shape shape13 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties13 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties16 = new NonVisualDrawingProperties(){ Id = (UInt32Value)2U, Name = "Title 1" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties13 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks12 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties13.Append(shapeLocks12);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties16 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape12 = new PlaceholderShape(){ Type = PlaceholderValues.Title };

            applicationNonVisualDrawingProperties16.Append(placeholderShape12);

            nonVisualShapeProperties13.Append(nonVisualDrawingProperties16);
            nonVisualShapeProperties13.Append(nonVisualShapeDrawingProperties13);
            nonVisualShapeProperties13.Append(applicationNonVisualDrawingProperties16);

            ShapeProperties shapeProperties13 = new ShapeProperties();

            A.Transform2D transform2D10 = new A.Transform2D();
            A.Offset offset13 = new A.Offset(){ X = 519907L, Y = 1424783L };
            A.Extents extents13 = new A.Extents(){ Cx = 6572250L, Cy = 2377281L };

            transform2D10.Append(offset13);
            transform2D10.Append(extents13);

            shapeProperties13.Append(transform2D10);

            TextBody textBody13 = new TextBody();
            A.BodyProperties bodyProperties13 = new A.BodyProperties(){ Anchor = A.TextAnchoringTypeValues.Bottom };

            A.ListStyle listStyle13 = new A.ListStyle();

            A.Level1ParagraphProperties level1ParagraphProperties10 = new A.Level1ParagraphProperties();
            A.DefaultRunProperties defaultRunProperties43 = new A.DefaultRunProperties(){ FontSize = 5000 };

            level1ParagraphProperties10.Append(defaultRunProperties43);

            listStyle13.Append(level1ParagraphProperties10);

            A.Paragraph paragraph21 = new A.Paragraph();

            A.Run run20 = new A.Run();
            A.RunProperties runProperties24 = new A.RunProperties(){ Language = "en-GB" };
            A.Text text24 = new A.Text();
            text24.Text = "Click to edit Master title style";

            run20.Append(runProperties24);
            run20.Append(text24);
            A.EndParagraphRunProperties endParagraphRunProperties12 = new A.EndParagraphRunProperties(){ Language = "en-US", Dirty = false };

            paragraph21.Append(run20);
            paragraph21.Append(endParagraphRunProperties12);

            textBody13.Append(bodyProperties13);
            textBody13.Append(listStyle13);
            textBody13.Append(paragraph21);

            shape13.Append(nonVisualShapeProperties13);
            shape13.Append(shapeProperties13);
            shape13.Append(textBody13);

            Shape shape14 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties14 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties17 = new NonVisualDrawingProperties(){ Id = (UInt32Value)3U, Name = "Text Placeholder 2" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties14 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks13 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties14.Append(shapeLocks13);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties17 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape13 = new PlaceholderShape(){ Type = PlaceholderValues.Body, Index = (UInt32Value)1U };

            applicationNonVisualDrawingProperties17.Append(placeholderShape13);

            nonVisualShapeProperties14.Append(nonVisualDrawingProperties17);
            nonVisualShapeProperties14.Append(nonVisualShapeDrawingProperties14);
            nonVisualShapeProperties14.Append(applicationNonVisualDrawingProperties17);

            ShapeProperties shapeProperties14 = new ShapeProperties();

            A.Transform2D transform2D11 = new A.Transform2D();
            A.Offset offset14 = new A.Offset(){ X = 519907L, Y = 3824554L };
            A.Extents extents14 = new A.Extents(){ Cx = 6572250L, Cy = 1250156L };

            transform2D11.Append(offset14);
            transform2D11.Append(extents14);

            shapeProperties14.Append(transform2D11);

            TextBody textBody14 = new TextBody();
            A.BodyProperties bodyProperties14 = new A.BodyProperties();

            A.ListStyle listStyle14 = new A.ListStyle();

            A.Level1ParagraphProperties level1ParagraphProperties11 = new A.Level1ParagraphProperties(){ LeftMargin = 0, Indent = 0 };
            A.NoBullet noBullet11 = new A.NoBullet();

            A.DefaultRunProperties defaultRunProperties44 = new A.DefaultRunProperties(){ FontSize = 2000 };

            A.SolidFill solidFill31 = new A.SolidFill();
            A.SchemeColor schemeColor47 = new A.SchemeColor(){ Val = A.SchemeColorValues.Text1 };

            solidFill31.Append(schemeColor47);

            defaultRunProperties44.Append(solidFill31);

            level1ParagraphProperties11.Append(noBullet11);
            level1ParagraphProperties11.Append(defaultRunProperties44);

            A.Level2ParagraphProperties level2ParagraphProperties5 = new A.Level2ParagraphProperties(){ LeftMargin = 380985, Indent = 0 };
            A.NoBullet noBullet12 = new A.NoBullet();

            A.DefaultRunProperties defaultRunProperties45 = new A.DefaultRunProperties(){ FontSize = 1667 };

            A.SolidFill solidFill32 = new A.SolidFill();

            A.SchemeColor schemeColor48 = new A.SchemeColor(){ Val = A.SchemeColorValues.Text1 };
            A.Tint tint11 = new A.Tint(){ Val = 75000 };

            schemeColor48.Append(tint11);

            solidFill32.Append(schemeColor48);

            defaultRunProperties45.Append(solidFill32);

            level2ParagraphProperties5.Append(noBullet12);
            level2ParagraphProperties5.Append(defaultRunProperties45);

            A.Level3ParagraphProperties level3ParagraphProperties5 = new A.Level3ParagraphProperties(){ LeftMargin = 761970, Indent = 0 };
            A.NoBullet noBullet13 = new A.NoBullet();

            A.DefaultRunProperties defaultRunProperties46 = new A.DefaultRunProperties(){ FontSize = 1500 };

            A.SolidFill solidFill33 = new A.SolidFill();

            A.SchemeColor schemeColor49 = new A.SchemeColor(){ Val = A.SchemeColorValues.Text1 };
            A.Tint tint12 = new A.Tint(){ Val = 75000 };

            schemeColor49.Append(tint12);

            solidFill33.Append(schemeColor49);

            defaultRunProperties46.Append(solidFill33);

            level3ParagraphProperties5.Append(noBullet13);
            level3ParagraphProperties5.Append(defaultRunProperties46);

            A.Level4ParagraphProperties level4ParagraphProperties5 = new A.Level4ParagraphProperties(){ LeftMargin = 1142954, Indent = 0 };
            A.NoBullet noBullet14 = new A.NoBullet();

            A.DefaultRunProperties defaultRunProperties47 = new A.DefaultRunProperties(){ FontSize = 1333 };

            A.SolidFill solidFill34 = new A.SolidFill();

            A.SchemeColor schemeColor50 = new A.SchemeColor(){ Val = A.SchemeColorValues.Text1 };
            A.Tint tint13 = new A.Tint(){ Val = 75000 };

            schemeColor50.Append(tint13);

            solidFill34.Append(schemeColor50);

            defaultRunProperties47.Append(solidFill34);

            level4ParagraphProperties5.Append(noBullet14);
            level4ParagraphProperties5.Append(defaultRunProperties47);

            A.Level5ParagraphProperties level5ParagraphProperties5 = new A.Level5ParagraphProperties(){ LeftMargin = 1523939, Indent = 0 };
            A.NoBullet noBullet15 = new A.NoBullet();

            A.DefaultRunProperties defaultRunProperties48 = new A.DefaultRunProperties(){ FontSize = 1333 };

            A.SolidFill solidFill35 = new A.SolidFill();

            A.SchemeColor schemeColor51 = new A.SchemeColor(){ Val = A.SchemeColorValues.Text1 };
            A.Tint tint14 = new A.Tint(){ Val = 75000 };

            schemeColor51.Append(tint14);

            solidFill35.Append(schemeColor51);

            defaultRunProperties48.Append(solidFill35);

            level5ParagraphProperties5.Append(noBullet15);
            level5ParagraphProperties5.Append(defaultRunProperties48);

            A.Level6ParagraphProperties level6ParagraphProperties5 = new A.Level6ParagraphProperties(){ LeftMargin = 1904924, Indent = 0 };
            A.NoBullet noBullet16 = new A.NoBullet();

            A.DefaultRunProperties defaultRunProperties49 = new A.DefaultRunProperties(){ FontSize = 1333 };

            A.SolidFill solidFill36 = new A.SolidFill();

            A.SchemeColor schemeColor52 = new A.SchemeColor(){ Val = A.SchemeColorValues.Text1 };
            A.Tint tint15 = new A.Tint(){ Val = 75000 };

            schemeColor52.Append(tint15);

            solidFill36.Append(schemeColor52);

            defaultRunProperties49.Append(solidFill36);

            level6ParagraphProperties5.Append(noBullet16);
            level6ParagraphProperties5.Append(defaultRunProperties49);

            A.Level7ParagraphProperties level7ParagraphProperties5 = new A.Level7ParagraphProperties(){ LeftMargin = 2285909, Indent = 0 };
            A.NoBullet noBullet17 = new A.NoBullet();

            A.DefaultRunProperties defaultRunProperties50 = new A.DefaultRunProperties(){ FontSize = 1333 };

            A.SolidFill solidFill37 = new A.SolidFill();

            A.SchemeColor schemeColor53 = new A.SchemeColor(){ Val = A.SchemeColorValues.Text1 };
            A.Tint tint16 = new A.Tint(){ Val = 75000 };

            schemeColor53.Append(tint16);

            solidFill37.Append(schemeColor53);

            defaultRunProperties50.Append(solidFill37);

            level7ParagraphProperties5.Append(noBullet17);
            level7ParagraphProperties5.Append(defaultRunProperties50);

            A.Level8ParagraphProperties level8ParagraphProperties5 = new A.Level8ParagraphProperties(){ LeftMargin = 2666893, Indent = 0 };
            A.NoBullet noBullet18 = new A.NoBullet();

            A.DefaultRunProperties defaultRunProperties51 = new A.DefaultRunProperties(){ FontSize = 1333 };

            A.SolidFill solidFill38 = new A.SolidFill();

            A.SchemeColor schemeColor54 = new A.SchemeColor(){ Val = A.SchemeColorValues.Text1 };
            A.Tint tint17 = new A.Tint(){ Val = 75000 };

            schemeColor54.Append(tint17);

            solidFill38.Append(schemeColor54);

            defaultRunProperties51.Append(solidFill38);

            level8ParagraphProperties5.Append(noBullet18);
            level8ParagraphProperties5.Append(defaultRunProperties51);

            A.Level9ParagraphProperties level9ParagraphProperties5 = new A.Level9ParagraphProperties(){ LeftMargin = 3047878, Indent = 0 };
            A.NoBullet noBullet19 = new A.NoBullet();

            A.DefaultRunProperties defaultRunProperties52 = new A.DefaultRunProperties(){ FontSize = 1333 };

            A.SolidFill solidFill39 = new A.SolidFill();

            A.SchemeColor schemeColor55 = new A.SchemeColor(){ Val = A.SchemeColorValues.Text1 };
            A.Tint tint18 = new A.Tint(){ Val = 75000 };

            schemeColor55.Append(tint18);

            solidFill39.Append(schemeColor55);

            defaultRunProperties52.Append(solidFill39);

            level9ParagraphProperties5.Append(noBullet19);
            level9ParagraphProperties5.Append(defaultRunProperties52);

            listStyle14.Append(level1ParagraphProperties11);
            listStyle14.Append(level2ParagraphProperties5);
            listStyle14.Append(level3ParagraphProperties5);
            listStyle14.Append(level4ParagraphProperties5);
            listStyle14.Append(level5ParagraphProperties5);
            listStyle14.Append(level6ParagraphProperties5);
            listStyle14.Append(level7ParagraphProperties5);
            listStyle14.Append(level8ParagraphProperties5);
            listStyle14.Append(level9ParagraphProperties5);

            A.Paragraph paragraph22 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties15 = new A.ParagraphProperties(){ Level = 0 };

            A.Run run21 = new A.Run();
            A.RunProperties runProperties25 = new A.RunProperties(){ Language = "en-GB" };
            A.Text text25 = new A.Text();
            text25.Text = "Click to edit Master text styles";

            run21.Append(runProperties25);
            run21.Append(text25);

            paragraph22.Append(paragraphProperties15);
            paragraph22.Append(run21);

            textBody14.Append(bodyProperties14);
            textBody14.Append(listStyle14);
            textBody14.Append(paragraph22);

            shape14.Append(nonVisualShapeProperties14);
            shape14.Append(shapeProperties14);
            shape14.Append(textBody14);

            Shape shape15 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties15 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties18 = new NonVisualDrawingProperties(){ Id = (UInt32Value)4U, Name = "Date Placeholder 3" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties15 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks14 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties15.Append(shapeLocks14);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties18 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape14 = new PlaceholderShape(){ Type = PlaceholderValues.DateAndTime, Size = PlaceholderSizeValues.Half, Index = (UInt32Value)10U };

            applicationNonVisualDrawingProperties18.Append(placeholderShape14);

            nonVisualShapeProperties15.Append(nonVisualDrawingProperties18);
            nonVisualShapeProperties15.Append(nonVisualShapeDrawingProperties15);
            nonVisualShapeProperties15.Append(applicationNonVisualDrawingProperties18);
            ShapeProperties shapeProperties15 = new ShapeProperties();

            TextBody textBody15 = new TextBody();
            A.BodyProperties bodyProperties15 = new A.BodyProperties();
            A.ListStyle listStyle15 = new A.ListStyle();

            A.Paragraph paragraph23 = new A.Paragraph();

            A.Field field5 = new A.Field(){ Id = "{D56BA034-03DE-42AB-8839-ACB7E58C6F65}", Type = "datetime1" };

            A.RunProperties runProperties26 = new A.RunProperties(){ Language = "en-US" };
            runProperties26.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            A.Text text26 = new A.Text();
            text26.Text = "1/17/2023";

            field5.Append(runProperties26);
            field5.Append(text26);
            A.EndParagraphRunProperties endParagraphRunProperties13 = new A.EndParagraphRunProperties(){ Language = "en-US", Dirty = false };

            paragraph23.Append(field5);
            paragraph23.Append(endParagraphRunProperties13);

            textBody15.Append(bodyProperties15);
            textBody15.Append(listStyle15);
            textBody15.Append(paragraph23);

            shape15.Append(nonVisualShapeProperties15);
            shape15.Append(shapeProperties15);
            shape15.Append(textBody15);

            Shape shape16 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties16 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties19 = new NonVisualDrawingProperties(){ Id = (UInt32Value)5U, Name = "Footer Placeholder 4" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties16 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks15 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties16.Append(shapeLocks15);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties19 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape15 = new PlaceholderShape(){ Type = PlaceholderValues.Footer, Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)11U };

            applicationNonVisualDrawingProperties19.Append(placeholderShape15);

            nonVisualShapeProperties16.Append(nonVisualDrawingProperties19);
            nonVisualShapeProperties16.Append(nonVisualShapeDrawingProperties16);
            nonVisualShapeProperties16.Append(applicationNonVisualDrawingProperties19);
            ShapeProperties shapeProperties16 = new ShapeProperties();

            TextBody textBody16 = new TextBody();
            A.BodyProperties bodyProperties16 = new A.BodyProperties();
            A.ListStyle listStyle16 = new A.ListStyle();

            A.Paragraph paragraph24 = new A.Paragraph();

            A.Run run22 = new A.Run();
            A.RunProperties runProperties27 = new A.RunProperties(){ Language = "en-US" };
            A.Text text27 = new A.Text();
            text27.Text = "Commercial & Workout Details";

            run22.Append(runProperties27);
            run22.Append(text27);
            A.EndParagraphRunProperties endParagraphRunProperties14 = new A.EndParagraphRunProperties(){ Language = "en-US", Dirty = false };

            paragraph24.Append(run22);
            paragraph24.Append(endParagraphRunProperties14);

            textBody16.Append(bodyProperties16);
            textBody16.Append(listStyle16);
            textBody16.Append(paragraph24);

            shape16.Append(nonVisualShapeProperties16);
            shape16.Append(shapeProperties16);
            shape16.Append(textBody16);

            Shape shape17 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties17 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties20 = new NonVisualDrawingProperties(){ Id = (UInt32Value)6U, Name = "Slide Number Placeholder 5" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties17 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks16 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties17.Append(shapeLocks16);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties20 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape16 = new PlaceholderShape(){ Type = PlaceholderValues.SlideNumber, Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)12U };

            applicationNonVisualDrawingProperties20.Append(placeholderShape16);

            nonVisualShapeProperties17.Append(nonVisualDrawingProperties20);
            nonVisualShapeProperties17.Append(nonVisualShapeDrawingProperties17);
            nonVisualShapeProperties17.Append(applicationNonVisualDrawingProperties20);
            ShapeProperties shapeProperties17 = new ShapeProperties();

            TextBody textBody17 = new TextBody();
            A.BodyProperties bodyProperties17 = new A.BodyProperties();
            A.ListStyle listStyle17 = new A.ListStyle();

            A.Paragraph paragraph25 = new A.Paragraph();

            A.Run run23 = new A.Run();

            A.RunProperties runProperties28 = new A.RunProperties(){ Language = "en-US" };

            A.SolidFill solidFill40 = new A.SolidFill();
            A.SchemeColor schemeColor56 = new A.SchemeColor(){ Val = A.SchemeColorValues.Accent2 };

            solidFill40.Append(schemeColor56);

            runProperties28.Append(solidFill40);
            A.Text text28 = new A.Text();
            text28.Text = "|";

            run23.Append(runProperties28);
            run23.Append(text28);

            A.Run run24 = new A.Run();
            A.RunProperties runProperties29 = new A.RunProperties(){ Language = "en-US" };
            A.Text text29 = new A.Text();
            text29.Text = "";

            run24.Append(runProperties29);
            run24.Append(text29);

            A.Field field6 = new A.Field(){ Id = "{E4F84C54-E2A4-46FF-B5B0-8F7A23C41D82}", Type = "slidenum" };

            A.RunProperties runProperties30 = new A.RunProperties(){ Language = "cs-CZ" };
            runProperties30.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            A.ParagraphProperties paragraphProperties16 = new A.ParagraphProperties();
            A.Text text30 = new A.Text();
            text30.Text = "‹#›";

            field6.Append(runProperties30);
            field6.Append(paragraphProperties16);
            field6.Append(text30);
            A.EndParagraphRunProperties endParagraphRunProperties15 = new A.EndParagraphRunProperties(){ Language = "cs-CZ", Dirty = false };

            paragraph25.Append(run23);
            paragraph25.Append(run24);
            paragraph25.Append(field6);
            paragraph25.Append(endParagraphRunProperties15);

            textBody17.Append(bodyProperties17);
            textBody17.Append(listStyle17);
            textBody17.Append(paragraph25);

            shape17.Append(nonVisualShapeProperties17);
            shape17.Append(shapeProperties17);
            shape17.Append(textBody17);

            shapeTree3.Append(nonVisualGroupShapeProperties3);
            shapeTree3.Append(groupShapeProperties3);
            shapeTree3.Append(shape13);
            shapeTree3.Append(shape14);
            shapeTree3.Append(shape15);
            shapeTree3.Append(shape16);
            shapeTree3.Append(shape17);

            CommonSlideDataExtensionList commonSlideDataExtensionList3 = new CommonSlideDataExtensionList();

            CommonSlideDataExtension commonSlideDataExtension3 = new CommonSlideDataExtension(){ Uri = "{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}" };

            P14.CreationId creationId3 = new P14.CreationId(){ Val = (UInt32Value)4023195697U };
            creationId3.AddNamespaceDeclaration("p14", "http://schemas.microsoft.com/office/powerpoint/2010/main");

            commonSlideDataExtension3.Append(creationId3);

            commonSlideDataExtensionList3.Append(commonSlideDataExtension3);

            commonSlideData3.Append(shapeTree3);
            commonSlideData3.Append(commonSlideDataExtensionList3);

            ColorMapOverride colorMapOverride2 = new ColorMapOverride();
            A.MasterColorMapping masterColorMapping2 = new A.MasterColorMapping();

            colorMapOverride2.Append(masterColorMapping2);

            slideLayout2.Append(commonSlideData3);
            slideLayout2.Append(colorMapOverride2);

            slideLayoutPart2.SlideLayout = slideLayout2;
        }

        // Generates content of slideLayoutPart3.
        private void GenerateSlideLayoutPart3Content(SlideLayoutPart slideLayoutPart3)
        {
            SlideLayout slideLayout3 = new SlideLayout(){ Type = SlideLayoutValues.Blank, Preserve = true };
            slideLayout3.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            slideLayout3.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            slideLayout3.AddNamespaceDeclaration("p", "http://schemas.openxmlformats.org/presentationml/2006/main");

            CommonSlideData commonSlideData4 = new CommonSlideData(){ Name = "Blank" };

            ShapeTree shapeTree4 = new ShapeTree();

            NonVisualGroupShapeProperties nonVisualGroupShapeProperties4 = new NonVisualGroupShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties21 = new NonVisualDrawingProperties(){ Id = (UInt32Value)1U, Name = "" };
            NonVisualGroupShapeDrawingProperties nonVisualGroupShapeDrawingProperties4 = new NonVisualGroupShapeDrawingProperties();
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties21 = new ApplicationNonVisualDrawingProperties();

            nonVisualGroupShapeProperties4.Append(nonVisualDrawingProperties21);
            nonVisualGroupShapeProperties4.Append(nonVisualGroupShapeDrawingProperties4);
            nonVisualGroupShapeProperties4.Append(applicationNonVisualDrawingProperties21);

            GroupShapeProperties groupShapeProperties4 = new GroupShapeProperties();

            A.TransformGroup transformGroup4 = new A.TransformGroup();
            A.Offset offset15 = new A.Offset(){ X = 0L, Y = 0L };
            A.Extents extents15 = new A.Extents(){ Cx = 0L, Cy = 0L };
            A.ChildOffset childOffset4 = new A.ChildOffset(){ X = 0L, Y = 0L };
            A.ChildExtents childExtents4 = new A.ChildExtents(){ Cx = 0L, Cy = 0L };

            transformGroup4.Append(offset15);
            transformGroup4.Append(extents15);
            transformGroup4.Append(childOffset4);
            transformGroup4.Append(childExtents4);

            groupShapeProperties4.Append(transformGroup4);

            Shape shape18 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties18 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties22 = new NonVisualDrawingProperties(){ Id = (UInt32Value)2U, Name = "Date Placeholder 1" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties18 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks17 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties18.Append(shapeLocks17);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties22 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape17 = new PlaceholderShape(){ Type = PlaceholderValues.DateAndTime, Size = PlaceholderSizeValues.Half, Index = (UInt32Value)10U };

            applicationNonVisualDrawingProperties22.Append(placeholderShape17);

            nonVisualShapeProperties18.Append(nonVisualDrawingProperties22);
            nonVisualShapeProperties18.Append(nonVisualShapeDrawingProperties18);
            nonVisualShapeProperties18.Append(applicationNonVisualDrawingProperties22);
            ShapeProperties shapeProperties18 = new ShapeProperties();

            TextBody textBody18 = new TextBody();
            A.BodyProperties bodyProperties18 = new A.BodyProperties();
            A.ListStyle listStyle18 = new A.ListStyle();

            A.Paragraph paragraph26 = new A.Paragraph();

            A.Field field7 = new A.Field(){ Id = "{5CF08B9E-79AA-4386-B5F3-83177A2FEE12}", Type = "datetime1" };

            A.RunProperties runProperties31 = new A.RunProperties(){ Language = "en-US" };
            runProperties31.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            A.Text text31 = new A.Text();
            text31.Text = "1/17/2023";

            field7.Append(runProperties31);
            field7.Append(text31);
            A.EndParagraphRunProperties endParagraphRunProperties16 = new A.EndParagraphRunProperties(){ Language = "en-US", Dirty = false };

            paragraph26.Append(field7);
            paragraph26.Append(endParagraphRunProperties16);

            textBody18.Append(bodyProperties18);
            textBody18.Append(listStyle18);
            textBody18.Append(paragraph26);

            shape18.Append(nonVisualShapeProperties18);
            shape18.Append(shapeProperties18);
            shape18.Append(textBody18);

            Shape shape19 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties19 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties23 = new NonVisualDrawingProperties(){ Id = (UInt32Value)3U, Name = "Footer Placeholder 2" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties19 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks18 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties19.Append(shapeLocks18);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties23 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape18 = new PlaceholderShape(){ Type = PlaceholderValues.Footer, Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)11U };

            applicationNonVisualDrawingProperties23.Append(placeholderShape18);

            nonVisualShapeProperties19.Append(nonVisualDrawingProperties23);
            nonVisualShapeProperties19.Append(nonVisualShapeDrawingProperties19);
            nonVisualShapeProperties19.Append(applicationNonVisualDrawingProperties23);
            ShapeProperties shapeProperties19 = new ShapeProperties();

            TextBody textBody19 = new TextBody();
            A.BodyProperties bodyProperties19 = new A.BodyProperties();
            A.ListStyle listStyle19 = new A.ListStyle();

            A.Paragraph paragraph27 = new A.Paragraph();

            A.Run run25 = new A.Run();
            A.RunProperties runProperties32 = new A.RunProperties(){ Language = "en-US" };
            A.Text text32 = new A.Text();
            text32.Text = "Commercial & Workout Details";

            run25.Append(runProperties32);
            run25.Append(text32);
            A.EndParagraphRunProperties endParagraphRunProperties17 = new A.EndParagraphRunProperties(){ Language = "en-US", Dirty = false };

            paragraph27.Append(run25);
            paragraph27.Append(endParagraphRunProperties17);

            textBody19.Append(bodyProperties19);
            textBody19.Append(listStyle19);
            textBody19.Append(paragraph27);

            shape19.Append(nonVisualShapeProperties19);
            shape19.Append(shapeProperties19);
            shape19.Append(textBody19);

            Shape shape20 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties20 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties24 = new NonVisualDrawingProperties(){ Id = (UInt32Value)4U, Name = "Slide Number Placeholder 3" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties20 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks19 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties20.Append(shapeLocks19);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties24 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape19 = new PlaceholderShape(){ Type = PlaceholderValues.SlideNumber, Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)12U };

            applicationNonVisualDrawingProperties24.Append(placeholderShape19);

            nonVisualShapeProperties20.Append(nonVisualDrawingProperties24);
            nonVisualShapeProperties20.Append(nonVisualShapeDrawingProperties20);
            nonVisualShapeProperties20.Append(applicationNonVisualDrawingProperties24);
            ShapeProperties shapeProperties20 = new ShapeProperties();

            TextBody textBody20 = new TextBody();
            A.BodyProperties bodyProperties20 = new A.BodyProperties();
            A.ListStyle listStyle20 = new A.ListStyle();

            A.Paragraph paragraph28 = new A.Paragraph();

            A.Run run26 = new A.Run();

            A.RunProperties runProperties33 = new A.RunProperties(){ Language = "en-US" };

            A.SolidFill solidFill41 = new A.SolidFill();
            A.SchemeColor schemeColor57 = new A.SchemeColor(){ Val = A.SchemeColorValues.Accent2 };

            solidFill41.Append(schemeColor57);

            runProperties33.Append(solidFill41);
            A.Text text33 = new A.Text();
            text33.Text = "|";

            run26.Append(runProperties33);
            run26.Append(text33);

            A.Run run27 = new A.Run();
            A.RunProperties runProperties34 = new A.RunProperties(){ Language = "en-US" };
            A.Text text34 = new A.Text();
            text34.Text = "";

            run27.Append(runProperties34);
            run27.Append(text34);

            A.Field field8 = new A.Field(){ Id = "{E4F84C54-E2A4-46FF-B5B0-8F7A23C41D82}", Type = "slidenum" };

            A.RunProperties runProperties35 = new A.RunProperties(){ Language = "cs-CZ" };
            runProperties35.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            A.ParagraphProperties paragraphProperties17 = new A.ParagraphProperties();
            A.Text text35 = new A.Text();
            text35.Text = "‹#›";

            field8.Append(runProperties35);
            field8.Append(paragraphProperties17);
            field8.Append(text35);
            A.EndParagraphRunProperties endParagraphRunProperties18 = new A.EndParagraphRunProperties(){ Language = "cs-CZ", Dirty = false };

            paragraph28.Append(run26);
            paragraph28.Append(run27);
            paragraph28.Append(field8);
            paragraph28.Append(endParagraphRunProperties18);

            textBody20.Append(bodyProperties20);
            textBody20.Append(listStyle20);
            textBody20.Append(paragraph28);

            shape20.Append(nonVisualShapeProperties20);
            shape20.Append(shapeProperties20);
            shape20.Append(textBody20);

            shapeTree4.Append(nonVisualGroupShapeProperties4);
            shapeTree4.Append(groupShapeProperties4);
            shapeTree4.Append(shape18);
            shapeTree4.Append(shape19);
            shapeTree4.Append(shape20);

            CommonSlideDataExtensionList commonSlideDataExtensionList4 = new CommonSlideDataExtensionList();

            CommonSlideDataExtension commonSlideDataExtension4 = new CommonSlideDataExtension(){ Uri = "{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}" };

            P14.CreationId creationId4 = new P14.CreationId(){ Val = (UInt32Value)871763921U };
            creationId4.AddNamespaceDeclaration("p14", "http://schemas.microsoft.com/office/powerpoint/2010/main");

            commonSlideDataExtension4.Append(creationId4);

            commonSlideDataExtensionList4.Append(commonSlideDataExtension4);

            commonSlideData4.Append(shapeTree4);
            commonSlideData4.Append(commonSlideDataExtensionList4);

            ColorMapOverride colorMapOverride3 = new ColorMapOverride();
            A.MasterColorMapping masterColorMapping3 = new A.MasterColorMapping();

            colorMapOverride3.Append(masterColorMapping3);

            slideLayout3.Append(commonSlideData4);
            slideLayout3.Append(colorMapOverride3);

            slideLayoutPart3.SlideLayout = slideLayout3;
        }

        // Generates content of slideLayoutPart4.
        private void GenerateSlideLayoutPart4Content(SlideLayoutPart slideLayoutPart4)
        {
            SlideLayout slideLayout4 = new SlideLayout(){ UserDrawn = true };
            slideLayout4.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            slideLayout4.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            slideLayout4.AddNamespaceDeclaration("p", "http://schemas.openxmlformats.org/presentationml/2006/main");

            CommonSlideData commonSlideData5 = new CommonSlideData(){ Name = "Titulka 1" };

            ShapeTree shapeTree5 = new ShapeTree();

            NonVisualGroupShapeProperties nonVisualGroupShapeProperties5 = new NonVisualGroupShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties25 = new NonVisualDrawingProperties(){ Id = (UInt32Value)1U, Name = "" };
            NonVisualGroupShapeDrawingProperties nonVisualGroupShapeDrawingProperties5 = new NonVisualGroupShapeDrawingProperties();
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties25 = new ApplicationNonVisualDrawingProperties();

            nonVisualGroupShapeProperties5.Append(nonVisualDrawingProperties25);
            nonVisualGroupShapeProperties5.Append(nonVisualGroupShapeDrawingProperties5);
            nonVisualGroupShapeProperties5.Append(applicationNonVisualDrawingProperties25);

            GroupShapeProperties groupShapeProperties5 = new GroupShapeProperties();

            A.TransformGroup transformGroup5 = new A.TransformGroup();
            A.Offset offset16 = new A.Offset(){ X = 0L, Y = 0L };
            A.Extents extents16 = new A.Extents(){ Cx = 0L, Cy = 0L };
            A.ChildOffset childOffset5 = new A.ChildOffset(){ X = 0L, Y = 0L };
            A.ChildExtents childExtents5 = new A.ChildExtents(){ Cx = 0L, Cy = 0L };

            transformGroup5.Append(offset16);
            transformGroup5.Append(extents16);
            transformGroup5.Append(childOffset5);
            transformGroup5.Append(childExtents5);

            groupShapeProperties5.Append(transformGroup5);

            Shape shape21 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties21 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties26 = new NonVisualDrawingProperties(){ Id = (UInt32Value)2U, Name = "Title 1" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties21 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks20 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties21.Append(shapeLocks20);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties26 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape20 = new PlaceholderShape(){ Type = PlaceholderValues.CenteredTitle };

            applicationNonVisualDrawingProperties26.Append(placeholderShape20);

            nonVisualShapeProperties21.Append(nonVisualDrawingProperties26);
            nonVisualShapeProperties21.Append(nonVisualShapeDrawingProperties21);
            nonVisualShapeProperties21.Append(applicationNonVisualDrawingProperties26);

            ShapeProperties shapeProperties21 = new ShapeProperties();

            A.Transform2D transform2D12 = new A.Transform2D();
            A.Offset offset17 = new A.Offset(){ X = 269875L, Y = 1881982L };
            A.Extents extents17 = new A.Extents(){ Cx = 5715000L, Cy = 1452034L };

            transform2D12.Append(offset17);
            transform2D12.Append(extents17);

            shapeProperties21.Append(transform2D12);

            TextBody textBody21 = new TextBody();

            A.BodyProperties bodyProperties21 = new A.BodyProperties(){ LeftInset = 0, TopInset = 0, RightInset = 0, BottomInset = 0, Anchor = A.TextAnchoringTypeValues.Top };
            A.NormalAutoFit normalAutoFit3 = new A.NormalAutoFit();

            bodyProperties21.Append(normalAutoFit3);

            A.ListStyle listStyle21 = new A.ListStyle();

            A.Level1ParagraphProperties level1ParagraphProperties12 = new A.Level1ParagraphProperties(){ Alignment = A.TextAlignmentTypeValues.Left };
            A.DefaultRunProperties defaultRunProperties53 = new A.DefaultRunProperties(){ FontSize = 3333 };

            level1ParagraphProperties12.Append(defaultRunProperties53);

            listStyle21.Append(level1ParagraphProperties12);

            A.Paragraph paragraph29 = new A.Paragraph();

            A.Run run28 = new A.Run();
            A.RunProperties runProperties36 = new A.RunProperties(){ Language = "en-US", Dirty = false };
            A.Text text36 = new A.Text();
            text36.Text = "Click to edit Master title style";

            run28.Append(runProperties36);
            run28.Append(text36);
            A.EndParagraphRunProperties endParagraphRunProperties19 = new A.EndParagraphRunProperties(){ Language = "cs-CZ", Dirty = false };

            paragraph29.Append(run28);
            paragraph29.Append(endParagraphRunProperties19);

            textBody21.Append(bodyProperties21);
            textBody21.Append(listStyle21);
            textBody21.Append(paragraph29);

            shape21.Append(nonVisualShapeProperties21);
            shape21.Append(shapeProperties21);
            shape21.Append(textBody21);

            Shape shape22 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties22 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties27 = new NonVisualDrawingProperties(){ Id = (UInt32Value)3U, Name = "Subtitle 2" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties22 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks21 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties22.Append(shapeLocks21);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties27 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape21 = new PlaceholderShape(){ Type = PlaceholderValues.SubTitle, Index = (UInt32Value)1U };

            applicationNonVisualDrawingProperties27.Append(placeholderShape21);

            nonVisualShapeProperties22.Append(nonVisualDrawingProperties27);
            nonVisualShapeProperties22.Append(nonVisualShapeDrawingProperties22);
            nonVisualShapeProperties22.Append(applicationNonVisualDrawingProperties27);

            ShapeProperties shapeProperties22 = new ShapeProperties();

            A.Transform2D transform2D13 = new A.Transform2D();
            A.Offset offset18 = new A.Offset(){ X = 269875L, Y = 3334016L };
            A.Extents extents18 = new A.Extents(){ Cx = 5715000L, Cy = 1379802L };

            transform2D13.Append(offset18);
            transform2D13.Append(extents18);

            shapeProperties22.Append(transform2D13);

            TextBody textBody22 = new TextBody();

            A.BodyProperties bodyProperties22 = new A.BodyProperties(){ LeftInset = 0, TopInset = 0, RightInset = 0, BottomInset = 0 };
            A.NormalAutoFit normalAutoFit4 = new A.NormalAutoFit();

            bodyProperties22.Append(normalAutoFit4);

            A.ListStyle listStyle22 = new A.ListStyle();

            A.Level1ParagraphProperties level1ParagraphProperties13 = new A.Level1ParagraphProperties(){ LeftMargin = 0, Indent = 0, Alignment = A.TextAlignmentTypeValues.Left };
            A.NoBullet noBullet20 = new A.NoBullet();

            A.DefaultRunProperties defaultRunProperties54 = new A.DefaultRunProperties(){ FontSize = 2333, Bold = true };

            A.SolidFill solidFill42 = new A.SolidFill();
            A.SchemeColor schemeColor58 = new A.SchemeColor(){ Val = A.SchemeColorValues.Accent2 };

            solidFill42.Append(schemeColor58);

            defaultRunProperties54.Append(solidFill42);

            level1ParagraphProperties13.Append(noBullet20);
            level1ParagraphProperties13.Append(defaultRunProperties54);

            A.Level2ParagraphProperties level2ParagraphProperties6 = new A.Level2ParagraphProperties(){ LeftMargin = 380985, Indent = 0, Alignment = A.TextAlignmentTypeValues.Center };
            A.NoBullet noBullet21 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties55 = new A.DefaultRunProperties(){ FontSize = 1667 };

            level2ParagraphProperties6.Append(noBullet21);
            level2ParagraphProperties6.Append(defaultRunProperties55);

            A.Level3ParagraphProperties level3ParagraphProperties6 = new A.Level3ParagraphProperties(){ LeftMargin = 761970, Indent = 0, Alignment = A.TextAlignmentTypeValues.Center };
            A.NoBullet noBullet22 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties56 = new A.DefaultRunProperties(){ FontSize = 1500 };

            level3ParagraphProperties6.Append(noBullet22);
            level3ParagraphProperties6.Append(defaultRunProperties56);

            A.Level4ParagraphProperties level4ParagraphProperties6 = new A.Level4ParagraphProperties(){ LeftMargin = 1142954, Indent = 0, Alignment = A.TextAlignmentTypeValues.Center };
            A.NoBullet noBullet23 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties57 = new A.DefaultRunProperties(){ FontSize = 1333 };

            level4ParagraphProperties6.Append(noBullet23);
            level4ParagraphProperties6.Append(defaultRunProperties57);

            A.Level5ParagraphProperties level5ParagraphProperties6 = new A.Level5ParagraphProperties(){ LeftMargin = 1523939, Indent = 0, Alignment = A.TextAlignmentTypeValues.Center };
            A.NoBullet noBullet24 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties58 = new A.DefaultRunProperties(){ FontSize = 1333 };

            level5ParagraphProperties6.Append(noBullet24);
            level5ParagraphProperties6.Append(defaultRunProperties58);

            A.Level6ParagraphProperties level6ParagraphProperties6 = new A.Level6ParagraphProperties(){ LeftMargin = 1904924, Indent = 0, Alignment = A.TextAlignmentTypeValues.Center };
            A.NoBullet noBullet25 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties59 = new A.DefaultRunProperties(){ FontSize = 1333 };

            level6ParagraphProperties6.Append(noBullet25);
            level6ParagraphProperties6.Append(defaultRunProperties59);

            A.Level7ParagraphProperties level7ParagraphProperties6 = new A.Level7ParagraphProperties(){ LeftMargin = 2285909, Indent = 0, Alignment = A.TextAlignmentTypeValues.Center };
            A.NoBullet noBullet26 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties60 = new A.DefaultRunProperties(){ FontSize = 1333 };

            level7ParagraphProperties6.Append(noBullet26);
            level7ParagraphProperties6.Append(defaultRunProperties60);

            A.Level8ParagraphProperties level8ParagraphProperties6 = new A.Level8ParagraphProperties(){ LeftMargin = 2666893, Indent = 0, Alignment = A.TextAlignmentTypeValues.Center };
            A.NoBullet noBullet27 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties61 = new A.DefaultRunProperties(){ FontSize = 1333 };

            level8ParagraphProperties6.Append(noBullet27);
            level8ParagraphProperties6.Append(defaultRunProperties61);

            A.Level9ParagraphProperties level9ParagraphProperties6 = new A.Level9ParagraphProperties(){ LeftMargin = 3047878, Indent = 0, Alignment = A.TextAlignmentTypeValues.Center };
            A.NoBullet noBullet28 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties62 = new A.DefaultRunProperties(){ FontSize = 1333 };

            level9ParagraphProperties6.Append(noBullet28);
            level9ParagraphProperties6.Append(defaultRunProperties62);

            listStyle22.Append(level1ParagraphProperties13);
            listStyle22.Append(level2ParagraphProperties6);
            listStyle22.Append(level3ParagraphProperties6);
            listStyle22.Append(level4ParagraphProperties6);
            listStyle22.Append(level5ParagraphProperties6);
            listStyle22.Append(level6ParagraphProperties6);
            listStyle22.Append(level7ParagraphProperties6);
            listStyle22.Append(level8ParagraphProperties6);
            listStyle22.Append(level9ParagraphProperties6);

            A.Paragraph paragraph30 = new A.Paragraph();

            A.Run run29 = new A.Run();
            A.RunProperties runProperties37 = new A.RunProperties(){ Language = "en-US", Dirty = false };
            A.Text text37 = new A.Text();
            text37.Text = "Click to edit Master subtitle style";

            run29.Append(runProperties37);
            run29.Append(text37);
            A.EndParagraphRunProperties endParagraphRunProperties20 = new A.EndParagraphRunProperties(){ Language = "cs-CZ", Dirty = false };

            paragraph30.Append(run29);
            paragraph30.Append(endParagraphRunProperties20);

            textBody22.Append(bodyProperties22);
            textBody22.Append(listStyle22);
            textBody22.Append(paragraph30);

            shape22.Append(nonVisualShapeProperties22);
            shape22.Append(shapeProperties22);
            shape22.Append(textBody22);

            Picture picture1 = new Picture();

            NonVisualPictureProperties nonVisualPictureProperties1 = new NonVisualPictureProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties28 = new NonVisualDrawingProperties(){ Id = (UInt32Value)7U, Name = "Picture 6" };

            NonVisualPictureDrawingProperties nonVisualPictureDrawingProperties1 = new NonVisualPictureDrawingProperties();
            A.PictureLocks pictureLocks1 = new A.PictureLocks(){ NoChangeAspect = true };

            nonVisualPictureDrawingProperties1.Append(pictureLocks1);
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties28 = new ApplicationNonVisualDrawingProperties(){ UserDrawn = true };

            nonVisualPictureProperties1.Append(nonVisualDrawingProperties28);
            nonVisualPictureProperties1.Append(nonVisualPictureDrawingProperties1);
            nonVisualPictureProperties1.Append(applicationNonVisualDrawingProperties28);

            BlipFill blipFill1 = new BlipFill(){ RotateWithShape = true };

            A.Blip blip1 = new A.Blip(){ Embed = "rId2", CompressionState = A.BlipCompressionValues.Print };

            A.BlipExtensionList blipExtensionList1 = new A.BlipExtensionList();

            A.BlipExtension blipExtension1 = new A.BlipExtension(){ Uri = "{28A0092B-C50C-407E-A947-70E740481C1C}" };

            A14.UseLocalDpi useLocalDpi1 = new A14.UseLocalDpi(){ Val = false };
            useLocalDpi1.AddNamespaceDeclaration("a14", "http://schemas.microsoft.com/office/drawing/2010/main");

            blipExtension1.Append(useLocalDpi1);

            blipExtensionList1.Append(blipExtension1);

            blip1.Append(blipExtensionList1);
            A.SourceRectangle sourceRectangle1 = new A.SourceRectangle(){ Left = 19025, Top = 49609 };
            A.Stretch stretch1 = new A.Stretch();

            blipFill1.Append(blip1);
            blipFill1.Append(sourceRectangle1);
            blipFill1.Append(stretch1);

            ShapeProperties shapeProperties23 = new ShapeProperties();

            A.Transform2D transform2D14 = new A.Transform2D();
            A.Offset offset19 = new A.Offset(){ X = 269875L, Y = 352451L };
            A.Extents extents19 = new A.Extents(){ Cx = 1991813L, Cy = 568673L };

            transform2D14.Append(offset19);
            transform2D14.Append(extents19);

            A.PresetGeometry presetGeometry7 = new A.PresetGeometry(){ Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList7 = new A.AdjustValueList();

            presetGeometry7.Append(adjustValueList7);

            shapeProperties23.Append(transform2D14);
            shapeProperties23.Append(presetGeometry7);

            picture1.Append(nonVisualPictureProperties1);
            picture1.Append(blipFill1);
            picture1.Append(shapeProperties23);

            Shape shape23 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties23 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties29 = new NonVisualDrawingProperties(){ Id = (UInt32Value)9U, Name = "Text Placeholder 8" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties23 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks22 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties23.Append(shapeLocks22);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties29 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape22 = new PlaceholderShape(){ Type = PlaceholderValues.Body, Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)10U };

            applicationNonVisualDrawingProperties29.Append(placeholderShape22);

            nonVisualShapeProperties23.Append(nonVisualDrawingProperties29);
            nonVisualShapeProperties23.Append(nonVisualShapeDrawingProperties23);
            nonVisualShapeProperties23.Append(applicationNonVisualDrawingProperties29);

            ShapeProperties shapeProperties24 = new ShapeProperties();

            A.Transform2D transform2D15 = new A.Transform2D();
            A.Offset offset20 = new A.Offset(){ X = 269875L, Y = 5058433L };
            A.Extents extents20 = new A.Extents(){ Cx = 5715000L, Cy = 327115L };

            transform2D15.Append(offset20);
            transform2D15.Append(extents20);

            shapeProperties24.Append(transform2D15);

            TextBody textBody23 = new TextBody();

            A.BodyProperties bodyProperties23 = new A.BodyProperties(){ LeftInset = 0, TopInset = 0, RightInset = 0, BottomInset = 0 };
            A.NormalAutoFit normalAutoFit5 = new A.NormalAutoFit();

            bodyProperties23.Append(normalAutoFit5);

            A.ListStyle listStyle23 = new A.ListStyle();

            A.Level1ParagraphProperties level1ParagraphProperties14 = new A.Level1ParagraphProperties(){ LeftMargin = 0, Indent = 0 };
            A.NoBullet noBullet29 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties63 = new A.DefaultRunProperties(){ FontSize = 1667 };

            level1ParagraphProperties14.Append(noBullet29);
            level1ParagraphProperties14.Append(defaultRunProperties63);

            A.Level2ParagraphProperties level2ParagraphProperties7 = new A.Level2ParagraphProperties(){ LeftMargin = 301613, Indent = 0 };
            A.NoBullet noBullet30 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties64 = new A.DefaultRunProperties();

            level2ParagraphProperties7.Append(noBullet30);
            level2ParagraphProperties7.Append(defaultRunProperties64);

            A.Level3ParagraphProperties level3ParagraphProperties7 = new A.Level3ParagraphProperties(){ LeftMargin = 0, Indent = 0 };
            A.BulletFont bulletFont10 = new A.BulletFont(){ Typeface = "Arial", Panose = "020B0604020202020204", PitchFamily = 34, CharacterSet = 0 };
            A.NoBullet noBullet31 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties65 = new A.DefaultRunProperties();

            level3ParagraphProperties7.Append(bulletFont10);
            level3ParagraphProperties7.Append(noBullet31);
            level3ParagraphProperties7.Append(defaultRunProperties65);

            A.Level4ParagraphProperties level4ParagraphProperties7 = new A.Level4ParagraphProperties(){ LeftMargin = 1142954, Indent = 0 };
            A.BulletFont bulletFont11 = new A.BulletFont(){ Typeface = "Arial", Panose = "020B0604020202020204", PitchFamily = 34, CharacterSet = 0 };
            A.NoBullet noBullet32 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties66 = new A.DefaultRunProperties();

            level4ParagraphProperties7.Append(bulletFont11);
            level4ParagraphProperties7.Append(noBullet32);
            level4ParagraphProperties7.Append(defaultRunProperties66);

            A.Level5ParagraphProperties level5ParagraphProperties7 = new A.Level5ParagraphProperties(){ LeftMargin = 1523939, Indent = 0 };
            A.NoBullet noBullet33 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties67 = new A.DefaultRunProperties();

            level5ParagraphProperties7.Append(noBullet33);
            level5ParagraphProperties7.Append(defaultRunProperties67);

            listStyle23.Append(level1ParagraphProperties14);
            listStyle23.Append(level2ParagraphProperties7);
            listStyle23.Append(level3ParagraphProperties7);
            listStyle23.Append(level4ParagraphProperties7);
            listStyle23.Append(level5ParagraphProperties7);

            A.Paragraph paragraph31 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties18 = new A.ParagraphProperties(){ Level = 0 };

            A.Run run30 = new A.Run();
            A.RunProperties runProperties38 = new A.RunProperties(){ Language = "en-US", Dirty = false };
            A.Text text38 = new A.Text();
            text38.Text = "Click to edit Master text styles";

            run30.Append(runProperties38);
            run30.Append(text38);

            paragraph31.Append(paragraphProperties18);
            paragraph31.Append(run30);

            textBody23.Append(bodyProperties23);
            textBody23.Append(listStyle23);
            textBody23.Append(paragraph31);

            shape23.Append(nonVisualShapeProperties23);
            shape23.Append(shapeProperties24);
            shape23.Append(textBody23);

            shapeTree5.Append(nonVisualGroupShapeProperties5);
            shapeTree5.Append(groupShapeProperties5);
            shapeTree5.Append(shape21);
            shapeTree5.Append(shape22);
            shapeTree5.Append(picture1);
            shapeTree5.Append(shape23);

            CommonSlideDataExtensionList commonSlideDataExtensionList5 = new CommonSlideDataExtensionList();

            CommonSlideDataExtension commonSlideDataExtension5 = new CommonSlideDataExtension(){ Uri = "{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}" };

            P14.CreationId creationId5 = new P14.CreationId(){ Val = (UInt32Value)3357994188U };
            creationId5.AddNamespaceDeclaration("p14", "http://schemas.microsoft.com/office/powerpoint/2010/main");

            commonSlideDataExtension5.Append(creationId5);

            commonSlideDataExtensionList5.Append(commonSlideDataExtension5);

            commonSlideData5.Append(shapeTree5);
            commonSlideData5.Append(commonSlideDataExtensionList5);

            ColorMapOverride colorMapOverride4 = new ColorMapOverride();
            A.MasterColorMapping masterColorMapping4 = new A.MasterColorMapping();

            colorMapOverride4.Append(masterColorMapping4);

            slideLayout4.Append(commonSlideData5);
            slideLayout4.Append(colorMapOverride4);

            slideLayoutPart4.SlideLayout = slideLayout4;
        }

        // Generates content of imagePart1.
        private void GenerateImagePart1Content(ImagePart imagePart1)
        {
            System.IO.Stream data = GetBinaryDataStream(imagePart1Data);
            imagePart1.FeedData(data);
            data.Close();
        }

        // Generates content of slideLayoutPart5.
        private void GenerateSlideLayoutPart5Content(SlideLayoutPart slideLayoutPart5)
        {
            SlideLayout slideLayout5 = new SlideLayout(){ Type = SlideLayoutValues.Object, Preserve = true };
            slideLayout5.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            slideLayout5.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            slideLayout5.AddNamespaceDeclaration("p", "http://schemas.openxmlformats.org/presentationml/2006/main");

            CommonSlideData commonSlideData6 = new CommonSlideData(){ Name = "Title and Content" };

            ShapeTree shapeTree6 = new ShapeTree();

            NonVisualGroupShapeProperties nonVisualGroupShapeProperties6 = new NonVisualGroupShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties30 = new NonVisualDrawingProperties(){ Id = (UInt32Value)1U, Name = "" };
            NonVisualGroupShapeDrawingProperties nonVisualGroupShapeDrawingProperties6 = new NonVisualGroupShapeDrawingProperties();
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties30 = new ApplicationNonVisualDrawingProperties();

            nonVisualGroupShapeProperties6.Append(nonVisualDrawingProperties30);
            nonVisualGroupShapeProperties6.Append(nonVisualGroupShapeDrawingProperties6);
            nonVisualGroupShapeProperties6.Append(applicationNonVisualDrawingProperties30);

            GroupShapeProperties groupShapeProperties6 = new GroupShapeProperties();

            A.TransformGroup transformGroup6 = new A.TransformGroup();
            A.Offset offset21 = new A.Offset(){ X = 0L, Y = 0L };
            A.Extents extents21 = new A.Extents(){ Cx = 0L, Cy = 0L };
            A.ChildOffset childOffset6 = new A.ChildOffset(){ X = 0L, Y = 0L };
            A.ChildExtents childExtents6 = new A.ChildExtents(){ Cx = 0L, Cy = 0L };

            transformGroup6.Append(offset21);
            transformGroup6.Append(extents21);
            transformGroup6.Append(childOffset6);
            transformGroup6.Append(childExtents6);

            groupShapeProperties6.Append(transformGroup6);

            Shape shape24 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties24 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties31 = new NonVisualDrawingProperties(){ Id = (UInt32Value)2U, Name = "Title 1" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties24 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks23 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties24.Append(shapeLocks23);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties31 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape23 = new PlaceholderShape(){ Type = PlaceholderValues.Title };

            applicationNonVisualDrawingProperties31.Append(placeholderShape23);

            nonVisualShapeProperties24.Append(nonVisualDrawingProperties31);
            nonVisualShapeProperties24.Append(nonVisualShapeDrawingProperties24);
            nonVisualShapeProperties24.Append(applicationNonVisualDrawingProperties31);
            ShapeProperties shapeProperties25 = new ShapeProperties();

            TextBody textBody24 = new TextBody();
            A.BodyProperties bodyProperties24 = new A.BodyProperties();
            A.ListStyle listStyle24 = new A.ListStyle();

            A.Paragraph paragraph32 = new A.Paragraph();

            A.Run run31 = new A.Run();
            A.RunProperties runProperties39 = new A.RunProperties(){ Language = "en-GB" };
            A.Text text39 = new A.Text();
            text39.Text = "Click to edit Master title style";

            run31.Append(runProperties39);
            run31.Append(text39);
            A.EndParagraphRunProperties endParagraphRunProperties21 = new A.EndParagraphRunProperties(){ Language = "en-US", Dirty = false };

            paragraph32.Append(run31);
            paragraph32.Append(endParagraphRunProperties21);

            textBody24.Append(bodyProperties24);
            textBody24.Append(listStyle24);
            textBody24.Append(paragraph32);

            shape24.Append(nonVisualShapeProperties24);
            shape24.Append(shapeProperties25);
            shape24.Append(textBody24);

            Shape shape25 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties25 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties32 = new NonVisualDrawingProperties(){ Id = (UInt32Value)3U, Name = "Content Placeholder 2" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties25 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks24 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties25.Append(shapeLocks24);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties32 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape24 = new PlaceholderShape(){ Index = (UInt32Value)1U };

            applicationNonVisualDrawingProperties32.Append(placeholderShape24);

            nonVisualShapeProperties25.Append(nonVisualDrawingProperties32);
            nonVisualShapeProperties25.Append(nonVisualShapeDrawingProperties25);
            nonVisualShapeProperties25.Append(applicationNonVisualDrawingProperties32);
            ShapeProperties shapeProperties26 = new ShapeProperties();

            TextBody textBody25 = new TextBody();
            A.BodyProperties bodyProperties25 = new A.BodyProperties();
            A.ListStyle listStyle25 = new A.ListStyle();

            A.Paragraph paragraph33 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties19 = new A.ParagraphProperties(){ Level = 0 };

            A.Run run32 = new A.Run();
            A.RunProperties runProperties40 = new A.RunProperties(){ Language = "en-GB" };
            A.Text text40 = new A.Text();
            text40.Text = "Click to edit Master text styles";

            run32.Append(runProperties40);
            run32.Append(text40);

            paragraph33.Append(paragraphProperties19);
            paragraph33.Append(run32);

            A.Paragraph paragraph34 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties20 = new A.ParagraphProperties(){ Level = 1 };

            A.Run run33 = new A.Run();
            A.RunProperties runProperties41 = new A.RunProperties(){ Language = "en-GB" };
            A.Text text41 = new A.Text();
            text41.Text = "Second level";

            run33.Append(runProperties41);
            run33.Append(text41);

            paragraph34.Append(paragraphProperties20);
            paragraph34.Append(run33);

            A.Paragraph paragraph35 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties21 = new A.ParagraphProperties(){ Level = 2 };

            A.Run run34 = new A.Run();
            A.RunProperties runProperties42 = new A.RunProperties(){ Language = "en-GB" };
            A.Text text42 = new A.Text();
            text42.Text = "Third level";

            run34.Append(runProperties42);
            run34.Append(text42);

            paragraph35.Append(paragraphProperties21);
            paragraph35.Append(run34);

            A.Paragraph paragraph36 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties22 = new A.ParagraphProperties(){ Level = 3 };

            A.Run run35 = new A.Run();
            A.RunProperties runProperties43 = new A.RunProperties(){ Language = "en-GB" };
            A.Text text43 = new A.Text();
            text43.Text = "Fourth level";

            run35.Append(runProperties43);
            run35.Append(text43);

            paragraph36.Append(paragraphProperties22);
            paragraph36.Append(run35);

            A.Paragraph paragraph37 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties23 = new A.ParagraphProperties(){ Level = 4 };

            A.Run run36 = new A.Run();
            A.RunProperties runProperties44 = new A.RunProperties(){ Language = "en-GB" };
            A.Text text44 = new A.Text();
            text44.Text = "Fifth level";

            run36.Append(runProperties44);
            run36.Append(text44);
            A.EndParagraphRunProperties endParagraphRunProperties22 = new A.EndParagraphRunProperties(){ Language = "en-US", Dirty = false };

            paragraph37.Append(paragraphProperties23);
            paragraph37.Append(run36);
            paragraph37.Append(endParagraphRunProperties22);

            textBody25.Append(bodyProperties25);
            textBody25.Append(listStyle25);
            textBody25.Append(paragraph33);
            textBody25.Append(paragraph34);
            textBody25.Append(paragraph35);
            textBody25.Append(paragraph36);
            textBody25.Append(paragraph37);

            shape25.Append(nonVisualShapeProperties25);
            shape25.Append(shapeProperties26);
            shape25.Append(textBody25);

            Shape shape26 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties26 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties33 = new NonVisualDrawingProperties(){ Id = (UInt32Value)4U, Name = "Date Placeholder 3" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties26 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks25 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties26.Append(shapeLocks25);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties33 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape25 = new PlaceholderShape(){ Type = PlaceholderValues.DateAndTime, Size = PlaceholderSizeValues.Half, Index = (UInt32Value)10U };

            applicationNonVisualDrawingProperties33.Append(placeholderShape25);

            nonVisualShapeProperties26.Append(nonVisualDrawingProperties33);
            nonVisualShapeProperties26.Append(nonVisualShapeDrawingProperties26);
            nonVisualShapeProperties26.Append(applicationNonVisualDrawingProperties33);
            ShapeProperties shapeProperties27 = new ShapeProperties();

            TextBody textBody26 = new TextBody();
            A.BodyProperties bodyProperties26 = new A.BodyProperties();
            A.ListStyle listStyle26 = new A.ListStyle();

            A.Paragraph paragraph38 = new A.Paragraph();

            A.Field field9 = new A.Field(){ Id = "{F6878D30-4F55-4AD9-9AEC-17531E549496}", Type = "datetime1" };

            A.RunProperties runProperties45 = new A.RunProperties(){ Language = "en-US" };
            runProperties45.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            A.Text text45 = new A.Text();
            text45.Text = "1/17/2023";

            field9.Append(runProperties45);
            field9.Append(text45);
            A.EndParagraphRunProperties endParagraphRunProperties23 = new A.EndParagraphRunProperties(){ Language = "en-US", Dirty = false };

            paragraph38.Append(field9);
            paragraph38.Append(endParagraphRunProperties23);

            textBody26.Append(bodyProperties26);
            textBody26.Append(listStyle26);
            textBody26.Append(paragraph38);

            shape26.Append(nonVisualShapeProperties26);
            shape26.Append(shapeProperties27);
            shape26.Append(textBody26);

            Shape shape27 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties27 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties34 = new NonVisualDrawingProperties(){ Id = (UInt32Value)5U, Name = "Footer Placeholder 4" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties27 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks26 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties27.Append(shapeLocks26);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties34 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape26 = new PlaceholderShape(){ Type = PlaceholderValues.Footer, Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)11U };

            applicationNonVisualDrawingProperties34.Append(placeholderShape26);

            nonVisualShapeProperties27.Append(nonVisualDrawingProperties34);
            nonVisualShapeProperties27.Append(nonVisualShapeDrawingProperties27);
            nonVisualShapeProperties27.Append(applicationNonVisualDrawingProperties34);
            ShapeProperties shapeProperties28 = new ShapeProperties();

            TextBody textBody27 = new TextBody();
            A.BodyProperties bodyProperties27 = new A.BodyProperties();
            A.ListStyle listStyle27 = new A.ListStyle();

            A.Paragraph paragraph39 = new A.Paragraph();

            A.Run run37 = new A.Run();
            A.RunProperties runProperties46 = new A.RunProperties(){ Language = "en-US" };
            A.Text text46 = new A.Text();
            text46.Text = "Commercial & Workout Details";

            run37.Append(runProperties46);
            run37.Append(text46);
            A.EndParagraphRunProperties endParagraphRunProperties24 = new A.EndParagraphRunProperties(){ Language = "en-US", Dirty = false };

            paragraph39.Append(run37);
            paragraph39.Append(endParagraphRunProperties24);

            textBody27.Append(bodyProperties27);
            textBody27.Append(listStyle27);
            textBody27.Append(paragraph39);

            shape27.Append(nonVisualShapeProperties27);
            shape27.Append(shapeProperties28);
            shape27.Append(textBody27);

            Shape shape28 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties28 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties35 = new NonVisualDrawingProperties(){ Id = (UInt32Value)6U, Name = "Slide Number Placeholder 5" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties28 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks27 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties28.Append(shapeLocks27);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties35 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape27 = new PlaceholderShape(){ Type = PlaceholderValues.SlideNumber, Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)12U };

            applicationNonVisualDrawingProperties35.Append(placeholderShape27);

            nonVisualShapeProperties28.Append(nonVisualDrawingProperties35);
            nonVisualShapeProperties28.Append(nonVisualShapeDrawingProperties28);
            nonVisualShapeProperties28.Append(applicationNonVisualDrawingProperties35);
            ShapeProperties shapeProperties29 = new ShapeProperties();

            TextBody textBody28 = new TextBody();
            A.BodyProperties bodyProperties28 = new A.BodyProperties();
            A.ListStyle listStyle28 = new A.ListStyle();

            A.Paragraph paragraph40 = new A.Paragraph();

            A.Run run38 = new A.Run();

            A.RunProperties runProperties47 = new A.RunProperties(){ Language = "en-US" };

            A.SolidFill solidFill43 = new A.SolidFill();
            A.SchemeColor schemeColor59 = new A.SchemeColor(){ Val = A.SchemeColorValues.Accent2 };

            solidFill43.Append(schemeColor59);

            runProperties47.Append(solidFill43);
            A.Text text47 = new A.Text();
            text47.Text = "|";

            run38.Append(runProperties47);
            run38.Append(text47);

            A.Run run39 = new A.Run();
            A.RunProperties runProperties48 = new A.RunProperties(){ Language = "en-US" };
            A.Text text48 = new A.Text();
            text48.Text = "";

            run39.Append(runProperties48);
            run39.Append(text48);

            A.Field field10 = new A.Field(){ Id = "{E4F84C54-E2A4-46FF-B5B0-8F7A23C41D82}", Type = "slidenum" };

            A.RunProperties runProperties49 = new A.RunProperties(){ Language = "cs-CZ" };
            runProperties49.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            A.ParagraphProperties paragraphProperties24 = new A.ParagraphProperties();
            A.Text text49 = new A.Text();
            text49.Text = "‹#›";

            field10.Append(runProperties49);
            field10.Append(paragraphProperties24);
            field10.Append(text49);
            A.EndParagraphRunProperties endParagraphRunProperties25 = new A.EndParagraphRunProperties(){ Language = "cs-CZ", Dirty = false };

            paragraph40.Append(run38);
            paragraph40.Append(run39);
            paragraph40.Append(field10);
            paragraph40.Append(endParagraphRunProperties25);

            textBody28.Append(bodyProperties28);
            textBody28.Append(listStyle28);
            textBody28.Append(paragraph40);

            shape28.Append(nonVisualShapeProperties28);
            shape28.Append(shapeProperties29);
            shape28.Append(textBody28);

            shapeTree6.Append(nonVisualGroupShapeProperties6);
            shapeTree6.Append(groupShapeProperties6);
            shapeTree6.Append(shape24);
            shapeTree6.Append(shape25);
            shapeTree6.Append(shape26);
            shapeTree6.Append(shape27);
            shapeTree6.Append(shape28);

            CommonSlideDataExtensionList commonSlideDataExtensionList6 = new CommonSlideDataExtensionList();

            CommonSlideDataExtension commonSlideDataExtension6 = new CommonSlideDataExtension(){ Uri = "{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}" };

            P14.CreationId creationId6 = new P14.CreationId(){ Val = (UInt32Value)3034527331U };
            creationId6.AddNamespaceDeclaration("p14", "http://schemas.microsoft.com/office/powerpoint/2010/main");

            commonSlideDataExtension6.Append(creationId6);

            commonSlideDataExtensionList6.Append(commonSlideDataExtension6);

            commonSlideData6.Append(shapeTree6);
            commonSlideData6.Append(commonSlideDataExtensionList6);

            ColorMapOverride colorMapOverride5 = new ColorMapOverride();
            A.MasterColorMapping masterColorMapping5 = new A.MasterColorMapping();

            colorMapOverride5.Append(masterColorMapping5);

            slideLayout5.Append(commonSlideData6);
            slideLayout5.Append(colorMapOverride5);

            slideLayoutPart5.SlideLayout = slideLayout5;
        }

        // Generates content of slideLayoutPart6.
        private void GenerateSlideLayoutPart6Content(SlideLayoutPart slideLayoutPart6)
        {
            SlideLayout slideLayout6 = new SlideLayout(){ Type = SlideLayoutValues.Title, Preserve = true };
            slideLayout6.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            slideLayout6.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            slideLayout6.AddNamespaceDeclaration("p", "http://schemas.openxmlformats.org/presentationml/2006/main");

            CommonSlideData commonSlideData7 = new CommonSlideData(){ Name = "Title Slide" };

            ShapeTree shapeTree7 = new ShapeTree();

            NonVisualGroupShapeProperties nonVisualGroupShapeProperties7 = new NonVisualGroupShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties36 = new NonVisualDrawingProperties(){ Id = (UInt32Value)1U, Name = "" };
            NonVisualGroupShapeDrawingProperties nonVisualGroupShapeDrawingProperties7 = new NonVisualGroupShapeDrawingProperties();
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties36 = new ApplicationNonVisualDrawingProperties();

            nonVisualGroupShapeProperties7.Append(nonVisualDrawingProperties36);
            nonVisualGroupShapeProperties7.Append(nonVisualGroupShapeDrawingProperties7);
            nonVisualGroupShapeProperties7.Append(applicationNonVisualDrawingProperties36);

            GroupShapeProperties groupShapeProperties7 = new GroupShapeProperties();

            A.TransformGroup transformGroup7 = new A.TransformGroup();
            A.Offset offset22 = new A.Offset(){ X = 0L, Y = 0L };
            A.Extents extents22 = new A.Extents(){ Cx = 0L, Cy = 0L };
            A.ChildOffset childOffset7 = new A.ChildOffset(){ X = 0L, Y = 0L };
            A.ChildExtents childExtents7 = new A.ChildExtents(){ Cx = 0L, Cy = 0L };

            transformGroup7.Append(offset22);
            transformGroup7.Append(extents22);
            transformGroup7.Append(childOffset7);
            transformGroup7.Append(childExtents7);

            groupShapeProperties7.Append(transformGroup7);

            Shape shape29 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties29 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties37 = new NonVisualDrawingProperties(){ Id = (UInt32Value)2U, Name = "Title 1" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties29 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks28 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties29.Append(shapeLocks28);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties37 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape28 = new PlaceholderShape(){ Type = PlaceholderValues.CenteredTitle };

            applicationNonVisualDrawingProperties37.Append(placeholderShape28);

            nonVisualShapeProperties29.Append(nonVisualDrawingProperties37);
            nonVisualShapeProperties29.Append(nonVisualShapeDrawingProperties29);
            nonVisualShapeProperties29.Append(applicationNonVisualDrawingProperties37);

            ShapeProperties shapeProperties30 = new ShapeProperties();

            A.Transform2D transform2D16 = new A.Transform2D();
            A.Offset offset23 = new A.Offset(){ X = 571500L, Y = 935302L };
            A.Extents extents23 = new A.Extents(){ Cx = 6477000L, Cy = 1989667L };

            transform2D16.Append(offset23);
            transform2D16.Append(extents23);

            shapeProperties30.Append(transform2D16);

            TextBody textBody29 = new TextBody();
            A.BodyProperties bodyProperties29 = new A.BodyProperties(){ Anchor = A.TextAnchoringTypeValues.Bottom };

            A.ListStyle listStyle29 = new A.ListStyle();

            A.Level1ParagraphProperties level1ParagraphProperties15 = new A.Level1ParagraphProperties(){ Alignment = A.TextAlignmentTypeValues.Center };
            A.DefaultRunProperties defaultRunProperties68 = new A.DefaultRunProperties(){ FontSize = 5000 };

            level1ParagraphProperties15.Append(defaultRunProperties68);

            listStyle29.Append(level1ParagraphProperties15);

            A.Paragraph paragraph41 = new A.Paragraph();

            A.Run run40 = new A.Run();
            A.RunProperties runProperties50 = new A.RunProperties(){ Language = "en-GB" };
            A.Text text50 = new A.Text();
            text50.Text = "Click to edit Master title style";

            run40.Append(runProperties50);
            run40.Append(text50);
            A.EndParagraphRunProperties endParagraphRunProperties26 = new A.EndParagraphRunProperties(){ Language = "en-US", Dirty = false };

            paragraph41.Append(run40);
            paragraph41.Append(endParagraphRunProperties26);

            textBody29.Append(bodyProperties29);
            textBody29.Append(listStyle29);
            textBody29.Append(paragraph41);

            shape29.Append(nonVisualShapeProperties29);
            shape29.Append(shapeProperties30);
            shape29.Append(textBody29);

            Shape shape30 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties30 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties38 = new NonVisualDrawingProperties(){ Id = (UInt32Value)3U, Name = "Subtitle 2" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties30 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks29 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties30.Append(shapeLocks29);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties38 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape29 = new PlaceholderShape(){ Type = PlaceholderValues.SubTitle, Index = (UInt32Value)1U };

            applicationNonVisualDrawingProperties38.Append(placeholderShape29);

            nonVisualShapeProperties30.Append(nonVisualDrawingProperties38);
            nonVisualShapeProperties30.Append(nonVisualShapeDrawingProperties30);
            nonVisualShapeProperties30.Append(applicationNonVisualDrawingProperties38);

            ShapeProperties shapeProperties31 = new ShapeProperties();

            A.Transform2D transform2D17 = new A.Transform2D();
            A.Offset offset24 = new A.Offset(){ X = 952500L, Y = 3001698L };
            A.Extents extents24 = new A.Extents(){ Cx = 5715000L, Cy = 1379802L };

            transform2D17.Append(offset24);
            transform2D17.Append(extents24);

            shapeProperties31.Append(transform2D17);

            TextBody textBody30 = new TextBody();
            A.BodyProperties bodyProperties30 = new A.BodyProperties();

            A.ListStyle listStyle30 = new A.ListStyle();

            A.Level1ParagraphProperties level1ParagraphProperties16 = new A.Level1ParagraphProperties(){ LeftMargin = 0, Indent = 0, Alignment = A.TextAlignmentTypeValues.Center };
            A.NoBullet noBullet34 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties69 = new A.DefaultRunProperties(){ FontSize = 2000 };

            level1ParagraphProperties16.Append(noBullet34);
            level1ParagraphProperties16.Append(defaultRunProperties69);

            A.Level2ParagraphProperties level2ParagraphProperties8 = new A.Level2ParagraphProperties(){ LeftMargin = 380985, Indent = 0, Alignment = A.TextAlignmentTypeValues.Center };
            A.NoBullet noBullet35 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties70 = new A.DefaultRunProperties(){ FontSize = 1667 };

            level2ParagraphProperties8.Append(noBullet35);
            level2ParagraphProperties8.Append(defaultRunProperties70);

            A.Level3ParagraphProperties level3ParagraphProperties8 = new A.Level3ParagraphProperties(){ LeftMargin = 761970, Indent = 0, Alignment = A.TextAlignmentTypeValues.Center };
            A.NoBullet noBullet36 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties71 = new A.DefaultRunProperties(){ FontSize = 1500 };

            level3ParagraphProperties8.Append(noBullet36);
            level3ParagraphProperties8.Append(defaultRunProperties71);

            A.Level4ParagraphProperties level4ParagraphProperties8 = new A.Level4ParagraphProperties(){ LeftMargin = 1142954, Indent = 0, Alignment = A.TextAlignmentTypeValues.Center };
            A.NoBullet noBullet37 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties72 = new A.DefaultRunProperties(){ FontSize = 1333 };

            level4ParagraphProperties8.Append(noBullet37);
            level4ParagraphProperties8.Append(defaultRunProperties72);

            A.Level5ParagraphProperties level5ParagraphProperties8 = new A.Level5ParagraphProperties(){ LeftMargin = 1523939, Indent = 0, Alignment = A.TextAlignmentTypeValues.Center };
            A.NoBullet noBullet38 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties73 = new A.DefaultRunProperties(){ FontSize = 1333 };

            level5ParagraphProperties8.Append(noBullet38);
            level5ParagraphProperties8.Append(defaultRunProperties73);

            A.Level6ParagraphProperties level6ParagraphProperties7 = new A.Level6ParagraphProperties(){ LeftMargin = 1904924, Indent = 0, Alignment = A.TextAlignmentTypeValues.Center };
            A.NoBullet noBullet39 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties74 = new A.DefaultRunProperties(){ FontSize = 1333 };

            level6ParagraphProperties7.Append(noBullet39);
            level6ParagraphProperties7.Append(defaultRunProperties74);

            A.Level7ParagraphProperties level7ParagraphProperties7 = new A.Level7ParagraphProperties(){ LeftMargin = 2285909, Indent = 0, Alignment = A.TextAlignmentTypeValues.Center };
            A.NoBullet noBullet40 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties75 = new A.DefaultRunProperties(){ FontSize = 1333 };

            level7ParagraphProperties7.Append(noBullet40);
            level7ParagraphProperties7.Append(defaultRunProperties75);

            A.Level8ParagraphProperties level8ParagraphProperties7 = new A.Level8ParagraphProperties(){ LeftMargin = 2666893, Indent = 0, Alignment = A.TextAlignmentTypeValues.Center };
            A.NoBullet noBullet41 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties76 = new A.DefaultRunProperties(){ FontSize = 1333 };

            level8ParagraphProperties7.Append(noBullet41);
            level8ParagraphProperties7.Append(defaultRunProperties76);

            A.Level9ParagraphProperties level9ParagraphProperties7 = new A.Level9ParagraphProperties(){ LeftMargin = 3047878, Indent = 0, Alignment = A.TextAlignmentTypeValues.Center };
            A.NoBullet noBullet42 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties77 = new A.DefaultRunProperties(){ FontSize = 1333 };

            level9ParagraphProperties7.Append(noBullet42);
            level9ParagraphProperties7.Append(defaultRunProperties77);

            listStyle30.Append(level1ParagraphProperties16);
            listStyle30.Append(level2ParagraphProperties8);
            listStyle30.Append(level3ParagraphProperties8);
            listStyle30.Append(level4ParagraphProperties8);
            listStyle30.Append(level5ParagraphProperties8);
            listStyle30.Append(level6ParagraphProperties7);
            listStyle30.Append(level7ParagraphProperties7);
            listStyle30.Append(level8ParagraphProperties7);
            listStyle30.Append(level9ParagraphProperties7);

            A.Paragraph paragraph42 = new A.Paragraph();

            A.Run run41 = new A.Run();
            A.RunProperties runProperties51 = new A.RunProperties(){ Language = "en-GB" };
            A.Text text51 = new A.Text();
            text51.Text = "Click to edit Master subtitle style";

            run41.Append(runProperties51);
            run41.Append(text51);
            A.EndParagraphRunProperties endParagraphRunProperties27 = new A.EndParagraphRunProperties(){ Language = "en-US", Dirty = false };

            paragraph42.Append(run41);
            paragraph42.Append(endParagraphRunProperties27);

            textBody30.Append(bodyProperties30);
            textBody30.Append(listStyle30);
            textBody30.Append(paragraph42);

            shape30.Append(nonVisualShapeProperties30);
            shape30.Append(shapeProperties31);
            shape30.Append(textBody30);

            Shape shape31 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties31 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties39 = new NonVisualDrawingProperties(){ Id = (UInt32Value)4U, Name = "Date Placeholder 3" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties31 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks30 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties31.Append(shapeLocks30);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties39 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape30 = new PlaceholderShape(){ Type = PlaceholderValues.DateAndTime, Size = PlaceholderSizeValues.Half, Index = (UInt32Value)10U };

            applicationNonVisualDrawingProperties39.Append(placeholderShape30);

            nonVisualShapeProperties31.Append(nonVisualDrawingProperties39);
            nonVisualShapeProperties31.Append(nonVisualShapeDrawingProperties31);
            nonVisualShapeProperties31.Append(applicationNonVisualDrawingProperties39);
            ShapeProperties shapeProperties32 = new ShapeProperties();

            TextBody textBody31 = new TextBody();
            A.BodyProperties bodyProperties31 = new A.BodyProperties();
            A.ListStyle listStyle31 = new A.ListStyle();

            A.Paragraph paragraph43 = new A.Paragraph();

            A.Field field11 = new A.Field(){ Id = "{57E77B1B-6C8E-4AE1-A524-B6D64012CA7E}", Type = "datetime1" };

            A.RunProperties runProperties52 = new A.RunProperties(){ Language = "en-US" };
            runProperties52.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            A.Text text52 = new A.Text();
            text52.Text = "1/17/2023";

            field11.Append(runProperties52);
            field11.Append(text52);
            A.EndParagraphRunProperties endParagraphRunProperties28 = new A.EndParagraphRunProperties(){ Language = "en-US", Dirty = false };

            paragraph43.Append(field11);
            paragraph43.Append(endParagraphRunProperties28);

            textBody31.Append(bodyProperties31);
            textBody31.Append(listStyle31);
            textBody31.Append(paragraph43);

            shape31.Append(nonVisualShapeProperties31);
            shape31.Append(shapeProperties32);
            shape31.Append(textBody31);

            Shape shape32 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties32 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties40 = new NonVisualDrawingProperties(){ Id = (UInt32Value)5U, Name = "Footer Placeholder 4" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties32 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks31 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties32.Append(shapeLocks31);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties40 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape31 = new PlaceholderShape(){ Type = PlaceholderValues.Footer, Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)11U };

            applicationNonVisualDrawingProperties40.Append(placeholderShape31);

            nonVisualShapeProperties32.Append(nonVisualDrawingProperties40);
            nonVisualShapeProperties32.Append(nonVisualShapeDrawingProperties32);
            nonVisualShapeProperties32.Append(applicationNonVisualDrawingProperties40);
            ShapeProperties shapeProperties33 = new ShapeProperties();

            TextBody textBody32 = new TextBody();
            A.BodyProperties bodyProperties32 = new A.BodyProperties();
            A.ListStyle listStyle32 = new A.ListStyle();

            A.Paragraph paragraph44 = new A.Paragraph();

            A.Run run42 = new A.Run();
            A.RunProperties runProperties53 = new A.RunProperties(){ Language = "en-US" };
            A.Text text53 = new A.Text();
            text53.Text = "Commercial & Workout Details";

            run42.Append(runProperties53);
            run42.Append(text53);
            A.EndParagraphRunProperties endParagraphRunProperties29 = new A.EndParagraphRunProperties(){ Language = "en-US", Dirty = false };

            paragraph44.Append(run42);
            paragraph44.Append(endParagraphRunProperties29);

            textBody32.Append(bodyProperties32);
            textBody32.Append(listStyle32);
            textBody32.Append(paragraph44);

            shape32.Append(nonVisualShapeProperties32);
            shape32.Append(shapeProperties33);
            shape32.Append(textBody32);

            Shape shape33 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties33 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties41 = new NonVisualDrawingProperties(){ Id = (UInt32Value)6U, Name = "Slide Number Placeholder 5" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties33 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks32 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties33.Append(shapeLocks32);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties41 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape32 = new PlaceholderShape(){ Type = PlaceholderValues.SlideNumber, Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)12U };

            applicationNonVisualDrawingProperties41.Append(placeholderShape32);

            nonVisualShapeProperties33.Append(nonVisualDrawingProperties41);
            nonVisualShapeProperties33.Append(nonVisualShapeDrawingProperties33);
            nonVisualShapeProperties33.Append(applicationNonVisualDrawingProperties41);
            ShapeProperties shapeProperties34 = new ShapeProperties();

            TextBody textBody33 = new TextBody();
            A.BodyProperties bodyProperties33 = new A.BodyProperties();
            A.ListStyle listStyle33 = new A.ListStyle();

            A.Paragraph paragraph45 = new A.Paragraph();

            A.Run run43 = new A.Run();

            A.RunProperties runProperties54 = new A.RunProperties(){ Language = "en-US" };

            A.SolidFill solidFill44 = new A.SolidFill();
            A.SchemeColor schemeColor60 = new A.SchemeColor(){ Val = A.SchemeColorValues.Accent2 };

            solidFill44.Append(schemeColor60);

            runProperties54.Append(solidFill44);
            A.Text text54 = new A.Text();
            text54.Text = "|";

            run43.Append(runProperties54);
            run43.Append(text54);

            A.Run run44 = new A.Run();
            A.RunProperties runProperties55 = new A.RunProperties(){ Language = "en-US" };
            A.Text text55 = new A.Text();
            text55.Text = "";

            run44.Append(runProperties55);
            run44.Append(text55);

            A.Field field12 = new A.Field(){ Id = "{E4F84C54-E2A4-46FF-B5B0-8F7A23C41D82}", Type = "slidenum" };

            A.RunProperties runProperties56 = new A.RunProperties(){ Language = "cs-CZ" };
            runProperties56.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            A.ParagraphProperties paragraphProperties25 = new A.ParagraphProperties();
            A.Text text56 = new A.Text();
            text56.Text = "‹#›";

            field12.Append(runProperties56);
            field12.Append(paragraphProperties25);
            field12.Append(text56);
            A.EndParagraphRunProperties endParagraphRunProperties30 = new A.EndParagraphRunProperties(){ Language = "cs-CZ", Dirty = false };

            paragraph45.Append(run43);
            paragraph45.Append(run44);
            paragraph45.Append(field12);
            paragraph45.Append(endParagraphRunProperties30);

            textBody33.Append(bodyProperties33);
            textBody33.Append(listStyle33);
            textBody33.Append(paragraph45);

            shape33.Append(nonVisualShapeProperties33);
            shape33.Append(shapeProperties34);
            shape33.Append(textBody33);

            shapeTree7.Append(nonVisualGroupShapeProperties7);
            shapeTree7.Append(groupShapeProperties7);
            shapeTree7.Append(shape29);
            shapeTree7.Append(shape30);
            shapeTree7.Append(shape31);
            shapeTree7.Append(shape32);
            shapeTree7.Append(shape33);

            CommonSlideDataExtensionList commonSlideDataExtensionList7 = new CommonSlideDataExtensionList();

            CommonSlideDataExtension commonSlideDataExtension7 = new CommonSlideDataExtension(){ Uri = "{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}" };

            P14.CreationId creationId7 = new P14.CreationId(){ Val = (UInt32Value)530374578U };
            creationId7.AddNamespaceDeclaration("p14", "http://schemas.microsoft.com/office/powerpoint/2010/main");

            commonSlideDataExtension7.Append(creationId7);

            commonSlideDataExtensionList7.Append(commonSlideDataExtension7);

            commonSlideData7.Append(shapeTree7);
            commonSlideData7.Append(commonSlideDataExtensionList7);

            ColorMapOverride colorMapOverride6 = new ColorMapOverride();
            A.MasterColorMapping masterColorMapping6 = new A.MasterColorMapping();

            colorMapOverride6.Append(masterColorMapping6);

            slideLayout6.Append(commonSlideData7);
            slideLayout6.Append(colorMapOverride6);

            slideLayoutPart6.SlideLayout = slideLayout6;
        }

        // Generates content of slideLayoutPart7.
        private void GenerateSlideLayoutPart7Content(SlideLayoutPart slideLayoutPart7)
        {
            SlideLayout slideLayout7 = new SlideLayout(){ Type = SlideLayoutValues.TitleOnly, Preserve = true };
            slideLayout7.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            slideLayout7.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            slideLayout7.AddNamespaceDeclaration("p", "http://schemas.openxmlformats.org/presentationml/2006/main");

            CommonSlideData commonSlideData8 = new CommonSlideData(){ Name = "Title Only" };

            ShapeTree shapeTree8 = new ShapeTree();

            NonVisualGroupShapeProperties nonVisualGroupShapeProperties8 = new NonVisualGroupShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties42 = new NonVisualDrawingProperties(){ Id = (UInt32Value)1U, Name = "" };
            NonVisualGroupShapeDrawingProperties nonVisualGroupShapeDrawingProperties8 = new NonVisualGroupShapeDrawingProperties();
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties42 = new ApplicationNonVisualDrawingProperties();

            nonVisualGroupShapeProperties8.Append(nonVisualDrawingProperties42);
            nonVisualGroupShapeProperties8.Append(nonVisualGroupShapeDrawingProperties8);
            nonVisualGroupShapeProperties8.Append(applicationNonVisualDrawingProperties42);

            GroupShapeProperties groupShapeProperties8 = new GroupShapeProperties();

            A.TransformGroup transformGroup8 = new A.TransformGroup();
            A.Offset offset25 = new A.Offset(){ X = 0L, Y = 0L };
            A.Extents extents25 = new A.Extents(){ Cx = 0L, Cy = 0L };
            A.ChildOffset childOffset8 = new A.ChildOffset(){ X = 0L, Y = 0L };
            A.ChildExtents childExtents8 = new A.ChildExtents(){ Cx = 0L, Cy = 0L };

            transformGroup8.Append(offset25);
            transformGroup8.Append(extents25);
            transformGroup8.Append(childOffset8);
            transformGroup8.Append(childExtents8);

            groupShapeProperties8.Append(transformGroup8);

            Shape shape34 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties34 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties43 = new NonVisualDrawingProperties(){ Id = (UInt32Value)2U, Name = "Title 1" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties34 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks33 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties34.Append(shapeLocks33);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties43 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape33 = new PlaceholderShape(){ Type = PlaceholderValues.Title };

            applicationNonVisualDrawingProperties43.Append(placeholderShape33);

            nonVisualShapeProperties34.Append(nonVisualDrawingProperties43);
            nonVisualShapeProperties34.Append(nonVisualShapeDrawingProperties34);
            nonVisualShapeProperties34.Append(applicationNonVisualDrawingProperties43);
            ShapeProperties shapeProperties35 = new ShapeProperties();

            TextBody textBody34 = new TextBody();
            A.BodyProperties bodyProperties34 = new A.BodyProperties();
            A.ListStyle listStyle34 = new A.ListStyle();

            A.Paragraph paragraph46 = new A.Paragraph();

            A.Run run45 = new A.Run();
            A.RunProperties runProperties57 = new A.RunProperties(){ Language = "en-GB" };
            A.Text text57 = new A.Text();
            text57.Text = "Click to edit Master title style";

            run45.Append(runProperties57);
            run45.Append(text57);
            A.EndParagraphRunProperties endParagraphRunProperties31 = new A.EndParagraphRunProperties(){ Language = "en-US", Dirty = false };

            paragraph46.Append(run45);
            paragraph46.Append(endParagraphRunProperties31);

            textBody34.Append(bodyProperties34);
            textBody34.Append(listStyle34);
            textBody34.Append(paragraph46);

            shape34.Append(nonVisualShapeProperties34);
            shape34.Append(shapeProperties35);
            shape34.Append(textBody34);

            Shape shape35 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties35 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties44 = new NonVisualDrawingProperties(){ Id = (UInt32Value)3U, Name = "Date Placeholder 2" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties35 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks34 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties35.Append(shapeLocks34);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties44 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape34 = new PlaceholderShape(){ Type = PlaceholderValues.DateAndTime, Size = PlaceholderSizeValues.Half, Index = (UInt32Value)10U };

            applicationNonVisualDrawingProperties44.Append(placeholderShape34);

            nonVisualShapeProperties35.Append(nonVisualDrawingProperties44);
            nonVisualShapeProperties35.Append(nonVisualShapeDrawingProperties35);
            nonVisualShapeProperties35.Append(applicationNonVisualDrawingProperties44);
            ShapeProperties shapeProperties36 = new ShapeProperties();

            TextBody textBody35 = new TextBody();
            A.BodyProperties bodyProperties35 = new A.BodyProperties();
            A.ListStyle listStyle35 = new A.ListStyle();

            A.Paragraph paragraph47 = new A.Paragraph();

            A.Field field13 = new A.Field(){ Id = "{08D17CBD-BC76-4D31-8A50-20AC3136F80D}", Type = "datetime1" };

            A.RunProperties runProperties58 = new A.RunProperties(){ Language = "en-US" };
            runProperties58.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            A.Text text58 = new A.Text();
            text58.Text = "1/17/2023";

            field13.Append(runProperties58);
            field13.Append(text58);
            A.EndParagraphRunProperties endParagraphRunProperties32 = new A.EndParagraphRunProperties(){ Language = "en-US", Dirty = false };

            paragraph47.Append(field13);
            paragraph47.Append(endParagraphRunProperties32);

            textBody35.Append(bodyProperties35);
            textBody35.Append(listStyle35);
            textBody35.Append(paragraph47);

            shape35.Append(nonVisualShapeProperties35);
            shape35.Append(shapeProperties36);
            shape35.Append(textBody35);

            Shape shape36 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties36 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties45 = new NonVisualDrawingProperties(){ Id = (UInt32Value)4U, Name = "Footer Placeholder 3" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties36 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks35 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties36.Append(shapeLocks35);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties45 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape35 = new PlaceholderShape(){ Type = PlaceholderValues.Footer, Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)11U };

            applicationNonVisualDrawingProperties45.Append(placeholderShape35);

            nonVisualShapeProperties36.Append(nonVisualDrawingProperties45);
            nonVisualShapeProperties36.Append(nonVisualShapeDrawingProperties36);
            nonVisualShapeProperties36.Append(applicationNonVisualDrawingProperties45);
            ShapeProperties shapeProperties37 = new ShapeProperties();

            TextBody textBody36 = new TextBody();
            A.BodyProperties bodyProperties36 = new A.BodyProperties();
            A.ListStyle listStyle36 = new A.ListStyle();

            A.Paragraph paragraph48 = new A.Paragraph();

            A.Run run46 = new A.Run();
            A.RunProperties runProperties59 = new A.RunProperties(){ Language = "en-US" };
            A.Text text59 = new A.Text();
            text59.Text = "Commercial & Workout Details";

            run46.Append(runProperties59);
            run46.Append(text59);
            A.EndParagraphRunProperties endParagraphRunProperties33 = new A.EndParagraphRunProperties(){ Language = "en-US", Dirty = false };

            paragraph48.Append(run46);
            paragraph48.Append(endParagraphRunProperties33);

            textBody36.Append(bodyProperties36);
            textBody36.Append(listStyle36);
            textBody36.Append(paragraph48);

            shape36.Append(nonVisualShapeProperties36);
            shape36.Append(shapeProperties37);
            shape36.Append(textBody36);

            Shape shape37 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties37 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties46 = new NonVisualDrawingProperties(){ Id = (UInt32Value)5U, Name = "Slide Number Placeholder 4" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties37 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks36 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties37.Append(shapeLocks36);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties46 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape36 = new PlaceholderShape(){ Type = PlaceholderValues.SlideNumber, Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)12U };

            applicationNonVisualDrawingProperties46.Append(placeholderShape36);

            nonVisualShapeProperties37.Append(nonVisualDrawingProperties46);
            nonVisualShapeProperties37.Append(nonVisualShapeDrawingProperties37);
            nonVisualShapeProperties37.Append(applicationNonVisualDrawingProperties46);
            ShapeProperties shapeProperties38 = new ShapeProperties();

            TextBody textBody37 = new TextBody();
            A.BodyProperties bodyProperties37 = new A.BodyProperties();
            A.ListStyle listStyle37 = new A.ListStyle();

            A.Paragraph paragraph49 = new A.Paragraph();

            A.Run run47 = new A.Run();

            A.RunProperties runProperties60 = new A.RunProperties(){ Language = "en-US" };

            A.SolidFill solidFill45 = new A.SolidFill();
            A.SchemeColor schemeColor61 = new A.SchemeColor(){ Val = A.SchemeColorValues.Accent2 };

            solidFill45.Append(schemeColor61);

            runProperties60.Append(solidFill45);
            A.Text text60 = new A.Text();
            text60.Text = "|";

            run47.Append(runProperties60);
            run47.Append(text60);

            A.Run run48 = new A.Run();
            A.RunProperties runProperties61 = new A.RunProperties(){ Language = "en-US" };
            A.Text text61 = new A.Text();
            text61.Text = "";

            run48.Append(runProperties61);
            run48.Append(text61);

            A.Field field14 = new A.Field(){ Id = "{E4F84C54-E2A4-46FF-B5B0-8F7A23C41D82}", Type = "slidenum" };

            A.RunProperties runProperties62 = new A.RunProperties(){ Language = "cs-CZ" };
            runProperties62.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            A.ParagraphProperties paragraphProperties26 = new A.ParagraphProperties();
            A.Text text62 = new A.Text();
            text62.Text = "‹#›";

            field14.Append(runProperties62);
            field14.Append(paragraphProperties26);
            field14.Append(text62);
            A.EndParagraphRunProperties endParagraphRunProperties34 = new A.EndParagraphRunProperties(){ Language = "cs-CZ", Dirty = false };

            paragraph49.Append(run47);
            paragraph49.Append(run48);
            paragraph49.Append(field14);
            paragraph49.Append(endParagraphRunProperties34);

            textBody37.Append(bodyProperties37);
            textBody37.Append(listStyle37);
            textBody37.Append(paragraph49);

            shape37.Append(nonVisualShapeProperties37);
            shape37.Append(shapeProperties38);
            shape37.Append(textBody37);

            shapeTree8.Append(nonVisualGroupShapeProperties8);
            shapeTree8.Append(groupShapeProperties8);
            shapeTree8.Append(shape34);
            shapeTree8.Append(shape35);
            shapeTree8.Append(shape36);
            shapeTree8.Append(shape37);

            CommonSlideDataExtensionList commonSlideDataExtensionList8 = new CommonSlideDataExtensionList();

            CommonSlideDataExtension commonSlideDataExtension8 = new CommonSlideDataExtension(){ Uri = "{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}" };

            P14.CreationId creationId8 = new P14.CreationId(){ Val = (UInt32Value)1288102168U };
            creationId8.AddNamespaceDeclaration("p14", "http://schemas.microsoft.com/office/powerpoint/2010/main");

            commonSlideDataExtension8.Append(creationId8);

            commonSlideDataExtensionList8.Append(commonSlideDataExtension8);

            commonSlideData8.Append(shapeTree8);
            commonSlideData8.Append(commonSlideDataExtensionList8);

            ColorMapOverride colorMapOverride7 = new ColorMapOverride();
            A.MasterColorMapping masterColorMapping7 = new A.MasterColorMapping();

            colorMapOverride7.Append(masterColorMapping7);

            slideLayout7.Append(commonSlideData8);
            slideLayout7.Append(colorMapOverride7);

            slideLayoutPart7.SlideLayout = slideLayout7;
        }

        // Generates content of slideLayoutPart8.
        private void GenerateSlideLayoutPart8Content(SlideLayoutPart slideLayoutPart8)
        {
            SlideLayout slideLayout8 = new SlideLayout(){ Type = SlideLayoutValues.VerticalTitleAndText, Preserve = true };
            slideLayout8.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            slideLayout8.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            slideLayout8.AddNamespaceDeclaration("p", "http://schemas.openxmlformats.org/presentationml/2006/main");

            CommonSlideData commonSlideData9 = new CommonSlideData(){ Name = "Vertical Title and Text" };

            ShapeTree shapeTree9 = new ShapeTree();

            NonVisualGroupShapeProperties nonVisualGroupShapeProperties9 = new NonVisualGroupShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties47 = new NonVisualDrawingProperties(){ Id = (UInt32Value)1U, Name = "" };
            NonVisualGroupShapeDrawingProperties nonVisualGroupShapeDrawingProperties9 = new NonVisualGroupShapeDrawingProperties();
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties47 = new ApplicationNonVisualDrawingProperties();

            nonVisualGroupShapeProperties9.Append(nonVisualDrawingProperties47);
            nonVisualGroupShapeProperties9.Append(nonVisualGroupShapeDrawingProperties9);
            nonVisualGroupShapeProperties9.Append(applicationNonVisualDrawingProperties47);

            GroupShapeProperties groupShapeProperties9 = new GroupShapeProperties();

            A.TransformGroup transformGroup9 = new A.TransformGroup();
            A.Offset offset26 = new A.Offset(){ X = 0L, Y = 0L };
            A.Extents extents26 = new A.Extents(){ Cx = 0L, Cy = 0L };
            A.ChildOffset childOffset9 = new A.ChildOffset(){ X = 0L, Y = 0L };
            A.ChildExtents childExtents9 = new A.ChildExtents(){ Cx = 0L, Cy = 0L };

            transformGroup9.Append(offset26);
            transformGroup9.Append(extents26);
            transformGroup9.Append(childOffset9);
            transformGroup9.Append(childExtents9);

            groupShapeProperties9.Append(transformGroup9);

            Shape shape38 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties38 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties48 = new NonVisualDrawingProperties(){ Id = (UInt32Value)2U, Name = "Vertical Title 1" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties38 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks37 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties38.Append(shapeLocks37);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties48 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape37 = new PlaceholderShape(){ Type = PlaceholderValues.Title, Orientation = DirectionValues.Vertical };

            applicationNonVisualDrawingProperties48.Append(placeholderShape37);

            nonVisualShapeProperties38.Append(nonVisualDrawingProperties48);
            nonVisualShapeProperties38.Append(nonVisualShapeDrawingProperties38);
            nonVisualShapeProperties38.Append(applicationNonVisualDrawingProperties48);

            ShapeProperties shapeProperties39 = new ShapeProperties();

            A.Transform2D transform2D18 = new A.Transform2D();
            A.Offset offset27 = new A.Offset(){ X = 5453063L, Y = 304271L };
            A.Extents extents27 = new A.Extents(){ Cx = 1643063L, Cy = 4843198L };

            transform2D18.Append(offset27);
            transform2D18.Append(extents27);

            shapeProperties39.Append(transform2D18);

            TextBody textBody38 = new TextBody();
            A.BodyProperties bodyProperties38 = new A.BodyProperties(){ Vertical = A.TextVerticalValues.EastAsianVetical };
            A.ListStyle listStyle38 = new A.ListStyle();

            A.Paragraph paragraph50 = new A.Paragraph();

            A.Run run49 = new A.Run();
            A.RunProperties runProperties63 = new A.RunProperties(){ Language = "en-GB" };
            A.Text text63 = new A.Text();
            text63.Text = "Click to edit Master title style";

            run49.Append(runProperties63);
            run49.Append(text63);
            A.EndParagraphRunProperties endParagraphRunProperties35 = new A.EndParagraphRunProperties(){ Language = "en-US", Dirty = false };

            paragraph50.Append(run49);
            paragraph50.Append(endParagraphRunProperties35);

            textBody38.Append(bodyProperties38);
            textBody38.Append(listStyle38);
            textBody38.Append(paragraph50);

            shape38.Append(nonVisualShapeProperties38);
            shape38.Append(shapeProperties39);
            shape38.Append(textBody38);

            Shape shape39 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties39 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties49 = new NonVisualDrawingProperties(){ Id = (UInt32Value)3U, Name = "Vertical Text Placeholder 2" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties39 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks38 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties39.Append(shapeLocks38);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties49 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape38 = new PlaceholderShape(){ Type = PlaceholderValues.Body, Orientation = DirectionValues.Vertical, Index = (UInt32Value)1U };

            applicationNonVisualDrawingProperties49.Append(placeholderShape38);

            nonVisualShapeProperties39.Append(nonVisualDrawingProperties49);
            nonVisualShapeProperties39.Append(nonVisualShapeDrawingProperties39);
            nonVisualShapeProperties39.Append(applicationNonVisualDrawingProperties49);

            ShapeProperties shapeProperties40 = new ShapeProperties();

            A.Transform2D transform2D19 = new A.Transform2D();
            A.Offset offset28 = new A.Offset(){ X = 523875L, Y = 304271L };
            A.Extents extents28 = new A.Extents(){ Cx = 4833938L, Cy = 4843198L };

            transform2D19.Append(offset28);
            transform2D19.Append(extents28);

            shapeProperties40.Append(transform2D19);

            TextBody textBody39 = new TextBody();
            A.BodyProperties bodyProperties39 = new A.BodyProperties(){ Vertical = A.TextVerticalValues.EastAsianVetical };
            A.ListStyle listStyle39 = new A.ListStyle();

            A.Paragraph paragraph51 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties27 = new A.ParagraphProperties(){ Level = 0 };

            A.Run run50 = new A.Run();
            A.RunProperties runProperties64 = new A.RunProperties(){ Language = "en-GB" };
            A.Text text64 = new A.Text();
            text64.Text = "Click to edit Master text styles";

            run50.Append(runProperties64);
            run50.Append(text64);

            paragraph51.Append(paragraphProperties27);
            paragraph51.Append(run50);

            A.Paragraph paragraph52 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties28 = new A.ParagraphProperties(){ Level = 1 };

            A.Run run51 = new A.Run();
            A.RunProperties runProperties65 = new A.RunProperties(){ Language = "en-GB" };
            A.Text text65 = new A.Text();
            text65.Text = "Second level";

            run51.Append(runProperties65);
            run51.Append(text65);

            paragraph52.Append(paragraphProperties28);
            paragraph52.Append(run51);

            A.Paragraph paragraph53 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties29 = new A.ParagraphProperties(){ Level = 2 };

            A.Run run52 = new A.Run();
            A.RunProperties runProperties66 = new A.RunProperties(){ Language = "en-GB" };
            A.Text text66 = new A.Text();
            text66.Text = "Third level";

            run52.Append(runProperties66);
            run52.Append(text66);

            paragraph53.Append(paragraphProperties29);
            paragraph53.Append(run52);

            A.Paragraph paragraph54 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties30 = new A.ParagraphProperties(){ Level = 3 };

            A.Run run53 = new A.Run();
            A.RunProperties runProperties67 = new A.RunProperties(){ Language = "en-GB" };
            A.Text text67 = new A.Text();
            text67.Text = "Fourth level";

            run53.Append(runProperties67);
            run53.Append(text67);

            paragraph54.Append(paragraphProperties30);
            paragraph54.Append(run53);

            A.Paragraph paragraph55 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties31 = new A.ParagraphProperties(){ Level = 4 };

            A.Run run54 = new A.Run();
            A.RunProperties runProperties68 = new A.RunProperties(){ Language = "en-GB" };
            A.Text text68 = new A.Text();
            text68.Text = "Fifth level";

            run54.Append(runProperties68);
            run54.Append(text68);
            A.EndParagraphRunProperties endParagraphRunProperties36 = new A.EndParagraphRunProperties(){ Language = "en-US", Dirty = false };

            paragraph55.Append(paragraphProperties31);
            paragraph55.Append(run54);
            paragraph55.Append(endParagraphRunProperties36);

            textBody39.Append(bodyProperties39);
            textBody39.Append(listStyle39);
            textBody39.Append(paragraph51);
            textBody39.Append(paragraph52);
            textBody39.Append(paragraph53);
            textBody39.Append(paragraph54);
            textBody39.Append(paragraph55);

            shape39.Append(nonVisualShapeProperties39);
            shape39.Append(shapeProperties40);
            shape39.Append(textBody39);

            Shape shape40 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties40 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties50 = new NonVisualDrawingProperties(){ Id = (UInt32Value)4U, Name = "Date Placeholder 3" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties40 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks39 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties40.Append(shapeLocks39);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties50 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape39 = new PlaceholderShape(){ Type = PlaceholderValues.DateAndTime, Size = PlaceholderSizeValues.Half, Index = (UInt32Value)10U };

            applicationNonVisualDrawingProperties50.Append(placeholderShape39);

            nonVisualShapeProperties40.Append(nonVisualDrawingProperties50);
            nonVisualShapeProperties40.Append(nonVisualShapeDrawingProperties40);
            nonVisualShapeProperties40.Append(applicationNonVisualDrawingProperties50);
            ShapeProperties shapeProperties41 = new ShapeProperties();

            TextBody textBody40 = new TextBody();
            A.BodyProperties bodyProperties40 = new A.BodyProperties();
            A.ListStyle listStyle40 = new A.ListStyle();

            A.Paragraph paragraph56 = new A.Paragraph();

            A.Field field15 = new A.Field(){ Id = "{202ACFD5-F2BE-4E87-B68F-0DE78D2A3182}", Type = "datetime1" };

            A.RunProperties runProperties69 = new A.RunProperties(){ Language = "en-US" };
            runProperties69.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            A.Text text69 = new A.Text();
            text69.Text = "1/17/2023";

            field15.Append(runProperties69);
            field15.Append(text69);
            A.EndParagraphRunProperties endParagraphRunProperties37 = new A.EndParagraphRunProperties(){ Language = "en-US", Dirty = false };

            paragraph56.Append(field15);
            paragraph56.Append(endParagraphRunProperties37);

            textBody40.Append(bodyProperties40);
            textBody40.Append(listStyle40);
            textBody40.Append(paragraph56);

            shape40.Append(nonVisualShapeProperties40);
            shape40.Append(shapeProperties41);
            shape40.Append(textBody40);

            Shape shape41 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties41 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties51 = new NonVisualDrawingProperties(){ Id = (UInt32Value)5U, Name = "Footer Placeholder 4" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties41 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks40 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties41.Append(shapeLocks40);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties51 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape40 = new PlaceholderShape(){ Type = PlaceholderValues.Footer, Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)11U };

            applicationNonVisualDrawingProperties51.Append(placeholderShape40);

            nonVisualShapeProperties41.Append(nonVisualDrawingProperties51);
            nonVisualShapeProperties41.Append(nonVisualShapeDrawingProperties41);
            nonVisualShapeProperties41.Append(applicationNonVisualDrawingProperties51);
            ShapeProperties shapeProperties42 = new ShapeProperties();

            TextBody textBody41 = new TextBody();
            A.BodyProperties bodyProperties41 = new A.BodyProperties();
            A.ListStyle listStyle41 = new A.ListStyle();

            A.Paragraph paragraph57 = new A.Paragraph();

            A.Run run55 = new A.Run();
            A.RunProperties runProperties70 = new A.RunProperties(){ Language = "en-US" };
            A.Text text70 = new A.Text();
            text70.Text = "Commercial & Workout Details";

            run55.Append(runProperties70);
            run55.Append(text70);
            A.EndParagraphRunProperties endParagraphRunProperties38 = new A.EndParagraphRunProperties(){ Language = "en-US", Dirty = false };

            paragraph57.Append(run55);
            paragraph57.Append(endParagraphRunProperties38);

            textBody41.Append(bodyProperties41);
            textBody41.Append(listStyle41);
            textBody41.Append(paragraph57);

            shape41.Append(nonVisualShapeProperties41);
            shape41.Append(shapeProperties42);
            shape41.Append(textBody41);

            Shape shape42 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties42 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties52 = new NonVisualDrawingProperties(){ Id = (UInt32Value)6U, Name = "Slide Number Placeholder 5" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties42 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks41 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties42.Append(shapeLocks41);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties52 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape41 = new PlaceholderShape(){ Type = PlaceholderValues.SlideNumber, Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)12U };

            applicationNonVisualDrawingProperties52.Append(placeholderShape41);

            nonVisualShapeProperties42.Append(nonVisualDrawingProperties52);
            nonVisualShapeProperties42.Append(nonVisualShapeDrawingProperties42);
            nonVisualShapeProperties42.Append(applicationNonVisualDrawingProperties52);
            ShapeProperties shapeProperties43 = new ShapeProperties();

            TextBody textBody42 = new TextBody();
            A.BodyProperties bodyProperties42 = new A.BodyProperties();
            A.ListStyle listStyle42 = new A.ListStyle();

            A.Paragraph paragraph58 = new A.Paragraph();

            A.Run run56 = new A.Run();

            A.RunProperties runProperties71 = new A.RunProperties(){ Language = "en-US" };

            A.SolidFill solidFill46 = new A.SolidFill();
            A.SchemeColor schemeColor62 = new A.SchemeColor(){ Val = A.SchemeColorValues.Accent2 };

            solidFill46.Append(schemeColor62);

            runProperties71.Append(solidFill46);
            A.Text text71 = new A.Text();
            text71.Text = "|";

            run56.Append(runProperties71);
            run56.Append(text71);

            A.Run run57 = new A.Run();
            A.RunProperties runProperties72 = new A.RunProperties(){ Language = "en-US" };
            A.Text text72 = new A.Text();
            text72.Text = "";

            run57.Append(runProperties72);
            run57.Append(text72);

            A.Field field16 = new A.Field(){ Id = "{E4F84C54-E2A4-46FF-B5B0-8F7A23C41D82}", Type = "slidenum" };

            A.RunProperties runProperties73 = new A.RunProperties(){ Language = "cs-CZ" };
            runProperties73.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            A.ParagraphProperties paragraphProperties32 = new A.ParagraphProperties();
            A.Text text73 = new A.Text();
            text73.Text = "‹#›";

            field16.Append(runProperties73);
            field16.Append(paragraphProperties32);
            field16.Append(text73);
            A.EndParagraphRunProperties endParagraphRunProperties39 = new A.EndParagraphRunProperties(){ Language = "cs-CZ", Dirty = false };

            paragraph58.Append(run56);
            paragraph58.Append(run57);
            paragraph58.Append(field16);
            paragraph58.Append(endParagraphRunProperties39);

            textBody42.Append(bodyProperties42);
            textBody42.Append(listStyle42);
            textBody42.Append(paragraph58);

            shape42.Append(nonVisualShapeProperties42);
            shape42.Append(shapeProperties43);
            shape42.Append(textBody42);

            shapeTree9.Append(nonVisualGroupShapeProperties9);
            shapeTree9.Append(groupShapeProperties9);
            shapeTree9.Append(shape38);
            shapeTree9.Append(shape39);
            shapeTree9.Append(shape40);
            shapeTree9.Append(shape41);
            shapeTree9.Append(shape42);

            CommonSlideDataExtensionList commonSlideDataExtensionList9 = new CommonSlideDataExtensionList();

            CommonSlideDataExtension commonSlideDataExtension9 = new CommonSlideDataExtension(){ Uri = "{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}" };

            P14.CreationId creationId9 = new P14.CreationId(){ Val = (UInt32Value)2735585002U };
            creationId9.AddNamespaceDeclaration("p14", "http://schemas.microsoft.com/office/powerpoint/2010/main");

            commonSlideDataExtension9.Append(creationId9);

            commonSlideDataExtensionList9.Append(commonSlideDataExtension9);

            commonSlideData9.Append(shapeTree9);
            commonSlideData9.Append(commonSlideDataExtensionList9);

            ColorMapOverride colorMapOverride8 = new ColorMapOverride();
            A.MasterColorMapping masterColorMapping8 = new A.MasterColorMapping();

            colorMapOverride8.Append(masterColorMapping8);

            slideLayout8.Append(commonSlideData9);
            slideLayout8.Append(colorMapOverride8);

            slideLayoutPart8.SlideLayout = slideLayout8;
        }

        // Generates content of slideLayoutPart9.
        private void GenerateSlideLayoutPart9Content(SlideLayoutPart slideLayoutPart9)
        {
            SlideLayout slideLayout9 = new SlideLayout(){ Type = SlideLayoutValues.TwoTextAndTwoObjects, Preserve = true };
            slideLayout9.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            slideLayout9.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            slideLayout9.AddNamespaceDeclaration("p", "http://schemas.openxmlformats.org/presentationml/2006/main");

            CommonSlideData commonSlideData10 = new CommonSlideData(){ Name = "Comparison" };

            ShapeTree shapeTree10 = new ShapeTree();

            NonVisualGroupShapeProperties nonVisualGroupShapeProperties10 = new NonVisualGroupShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties53 = new NonVisualDrawingProperties(){ Id = (UInt32Value)1U, Name = "" };
            NonVisualGroupShapeDrawingProperties nonVisualGroupShapeDrawingProperties10 = new NonVisualGroupShapeDrawingProperties();
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties53 = new ApplicationNonVisualDrawingProperties();

            nonVisualGroupShapeProperties10.Append(nonVisualDrawingProperties53);
            nonVisualGroupShapeProperties10.Append(nonVisualGroupShapeDrawingProperties10);
            nonVisualGroupShapeProperties10.Append(applicationNonVisualDrawingProperties53);

            GroupShapeProperties groupShapeProperties10 = new GroupShapeProperties();

            A.TransformGroup transformGroup10 = new A.TransformGroup();
            A.Offset offset29 = new A.Offset(){ X = 0L, Y = 0L };
            A.Extents extents29 = new A.Extents(){ Cx = 0L, Cy = 0L };
            A.ChildOffset childOffset10 = new A.ChildOffset(){ X = 0L, Y = 0L };
            A.ChildExtents childExtents10 = new A.ChildExtents(){ Cx = 0L, Cy = 0L };

            transformGroup10.Append(offset29);
            transformGroup10.Append(extents29);
            transformGroup10.Append(childOffset10);
            transformGroup10.Append(childExtents10);

            groupShapeProperties10.Append(transformGroup10);

            Shape shape43 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties43 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties54 = new NonVisualDrawingProperties(){ Id = (UInt32Value)2U, Name = "Title 1" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties43 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks42 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties43.Append(shapeLocks42);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties54 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape42 = new PlaceholderShape(){ Type = PlaceholderValues.Title };

            applicationNonVisualDrawingProperties54.Append(placeholderShape42);

            nonVisualShapeProperties43.Append(nonVisualDrawingProperties54);
            nonVisualShapeProperties43.Append(nonVisualShapeDrawingProperties43);
            nonVisualShapeProperties43.Append(applicationNonVisualDrawingProperties54);

            ShapeProperties shapeProperties44 = new ShapeProperties();

            A.Transform2D transform2D20 = new A.Transform2D();
            A.Offset offset30 = new A.Offset(){ X = 524867L, Y = 304272L };
            A.Extents extents30 = new A.Extents(){ Cx = 6572250L, Cy = 1104636L };

            transform2D20.Append(offset30);
            transform2D20.Append(extents30);

            shapeProperties44.Append(transform2D20);

            TextBody textBody43 = new TextBody();
            A.BodyProperties bodyProperties43 = new A.BodyProperties();
            A.ListStyle listStyle43 = new A.ListStyle();

            A.Paragraph paragraph59 = new A.Paragraph();

            A.Run run58 = new A.Run();
            A.RunProperties runProperties74 = new A.RunProperties(){ Language = "en-GB" };
            A.Text text74 = new A.Text();
            text74.Text = "Click to edit Master title style";

            run58.Append(runProperties74);
            run58.Append(text74);
            A.EndParagraphRunProperties endParagraphRunProperties40 = new A.EndParagraphRunProperties(){ Language = "en-US", Dirty = false };

            paragraph59.Append(run58);
            paragraph59.Append(endParagraphRunProperties40);

            textBody43.Append(bodyProperties43);
            textBody43.Append(listStyle43);
            textBody43.Append(paragraph59);

            shape43.Append(nonVisualShapeProperties43);
            shape43.Append(shapeProperties44);
            shape43.Append(textBody43);

            Shape shape44 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties44 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties55 = new NonVisualDrawingProperties(){ Id = (UInt32Value)3U, Name = "Text Placeholder 2" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties44 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks43 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties44.Append(shapeLocks43);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties55 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape43 = new PlaceholderShape(){ Type = PlaceholderValues.Body, Index = (UInt32Value)1U };

            applicationNonVisualDrawingProperties55.Append(placeholderShape43);

            nonVisualShapeProperties44.Append(nonVisualDrawingProperties55);
            nonVisualShapeProperties44.Append(nonVisualShapeDrawingProperties44);
            nonVisualShapeProperties44.Append(applicationNonVisualDrawingProperties55);

            ShapeProperties shapeProperties45 = new ShapeProperties();

            A.Transform2D transform2D21 = new A.Transform2D();
            A.Offset offset31 = new A.Offset(){ X = 524868L, Y = 1400969L };
            A.Extents extents31 = new A.Extents(){ Cx = 3223617L, Cy = 686593L };

            transform2D21.Append(offset31);
            transform2D21.Append(extents31);

            shapeProperties45.Append(transform2D21);

            TextBody textBody44 = new TextBody();
            A.BodyProperties bodyProperties44 = new A.BodyProperties(){ Anchor = A.TextAnchoringTypeValues.Bottom };

            A.ListStyle listStyle44 = new A.ListStyle();

            A.Level1ParagraphProperties level1ParagraphProperties17 = new A.Level1ParagraphProperties(){ LeftMargin = 0, Indent = 0 };
            A.NoBullet noBullet43 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties78 = new A.DefaultRunProperties(){ FontSize = 2000, Bold = true };

            level1ParagraphProperties17.Append(noBullet43);
            level1ParagraphProperties17.Append(defaultRunProperties78);

            A.Level2ParagraphProperties level2ParagraphProperties9 = new A.Level2ParagraphProperties(){ LeftMargin = 380985, Indent = 0 };
            A.NoBullet noBullet44 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties79 = new A.DefaultRunProperties(){ FontSize = 1667, Bold = true };

            level2ParagraphProperties9.Append(noBullet44);
            level2ParagraphProperties9.Append(defaultRunProperties79);

            A.Level3ParagraphProperties level3ParagraphProperties9 = new A.Level3ParagraphProperties(){ LeftMargin = 761970, Indent = 0 };
            A.NoBullet noBullet45 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties80 = new A.DefaultRunProperties(){ FontSize = 1500, Bold = true };

            level3ParagraphProperties9.Append(noBullet45);
            level3ParagraphProperties9.Append(defaultRunProperties80);

            A.Level4ParagraphProperties level4ParagraphProperties9 = new A.Level4ParagraphProperties(){ LeftMargin = 1142954, Indent = 0 };
            A.NoBullet noBullet46 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties81 = new A.DefaultRunProperties(){ FontSize = 1333, Bold = true };

            level4ParagraphProperties9.Append(noBullet46);
            level4ParagraphProperties9.Append(defaultRunProperties81);

            A.Level5ParagraphProperties level5ParagraphProperties9 = new A.Level5ParagraphProperties(){ LeftMargin = 1523939, Indent = 0 };
            A.NoBullet noBullet47 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties82 = new A.DefaultRunProperties(){ FontSize = 1333, Bold = true };

            level5ParagraphProperties9.Append(noBullet47);
            level5ParagraphProperties9.Append(defaultRunProperties82);

            A.Level6ParagraphProperties level6ParagraphProperties8 = new A.Level6ParagraphProperties(){ LeftMargin = 1904924, Indent = 0 };
            A.NoBullet noBullet48 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties83 = new A.DefaultRunProperties(){ FontSize = 1333, Bold = true };

            level6ParagraphProperties8.Append(noBullet48);
            level6ParagraphProperties8.Append(defaultRunProperties83);

            A.Level7ParagraphProperties level7ParagraphProperties8 = new A.Level7ParagraphProperties(){ LeftMargin = 2285909, Indent = 0 };
            A.NoBullet noBullet49 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties84 = new A.DefaultRunProperties(){ FontSize = 1333, Bold = true };

            level7ParagraphProperties8.Append(noBullet49);
            level7ParagraphProperties8.Append(defaultRunProperties84);

            A.Level8ParagraphProperties level8ParagraphProperties8 = new A.Level8ParagraphProperties(){ LeftMargin = 2666893, Indent = 0 };
            A.NoBullet noBullet50 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties85 = new A.DefaultRunProperties(){ FontSize = 1333, Bold = true };

            level8ParagraphProperties8.Append(noBullet50);
            level8ParagraphProperties8.Append(defaultRunProperties85);

            A.Level9ParagraphProperties level9ParagraphProperties8 = new A.Level9ParagraphProperties(){ LeftMargin = 3047878, Indent = 0 };
            A.NoBullet noBullet51 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties86 = new A.DefaultRunProperties(){ FontSize = 1333, Bold = true };

            level9ParagraphProperties8.Append(noBullet51);
            level9ParagraphProperties8.Append(defaultRunProperties86);

            listStyle44.Append(level1ParagraphProperties17);
            listStyle44.Append(level2ParagraphProperties9);
            listStyle44.Append(level3ParagraphProperties9);
            listStyle44.Append(level4ParagraphProperties9);
            listStyle44.Append(level5ParagraphProperties9);
            listStyle44.Append(level6ParagraphProperties8);
            listStyle44.Append(level7ParagraphProperties8);
            listStyle44.Append(level8ParagraphProperties8);
            listStyle44.Append(level9ParagraphProperties8);

            A.Paragraph paragraph60 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties33 = new A.ParagraphProperties(){ Level = 0 };

            A.Run run59 = new A.Run();
            A.RunProperties runProperties75 = new A.RunProperties(){ Language = "en-GB" };
            A.Text text75 = new A.Text();
            text75.Text = "Click to edit Master text styles";

            run59.Append(runProperties75);
            run59.Append(text75);

            paragraph60.Append(paragraphProperties33);
            paragraph60.Append(run59);

            textBody44.Append(bodyProperties44);
            textBody44.Append(listStyle44);
            textBody44.Append(paragraph60);

            shape44.Append(nonVisualShapeProperties44);
            shape44.Append(shapeProperties45);
            shape44.Append(textBody44);

            Shape shape45 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties45 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties56 = new NonVisualDrawingProperties(){ Id = (UInt32Value)4U, Name = "Content Placeholder 3" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties45 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks44 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties45.Append(shapeLocks44);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties56 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape44 = new PlaceholderShape(){ Size = PlaceholderSizeValues.Half, Index = (UInt32Value)2U };

            applicationNonVisualDrawingProperties56.Append(placeholderShape44);

            nonVisualShapeProperties45.Append(nonVisualDrawingProperties56);
            nonVisualShapeProperties45.Append(nonVisualShapeDrawingProperties45);
            nonVisualShapeProperties45.Append(applicationNonVisualDrawingProperties56);

            ShapeProperties shapeProperties46 = new ShapeProperties();

            A.Transform2D transform2D22 = new A.Transform2D();
            A.Offset offset32 = new A.Offset(){ X = 524868L, Y = 2087563L };
            A.Extents extents32 = new A.Extents(){ Cx = 3223617L, Cy = 3070490L };

            transform2D22.Append(offset32);
            transform2D22.Append(extents32);

            shapeProperties46.Append(transform2D22);

            TextBody textBody45 = new TextBody();
            A.BodyProperties bodyProperties45 = new A.BodyProperties();
            A.ListStyle listStyle45 = new A.ListStyle();

            A.Paragraph paragraph61 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties34 = new A.ParagraphProperties(){ Level = 0 };

            A.Run run60 = new A.Run();
            A.RunProperties runProperties76 = new A.RunProperties(){ Language = "en-GB" };
            A.Text text76 = new A.Text();
            text76.Text = "Click to edit Master text styles";

            run60.Append(runProperties76);
            run60.Append(text76);

            paragraph61.Append(paragraphProperties34);
            paragraph61.Append(run60);

            A.Paragraph paragraph62 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties35 = new A.ParagraphProperties(){ Level = 1 };

            A.Run run61 = new A.Run();
            A.RunProperties runProperties77 = new A.RunProperties(){ Language = "en-GB" };
            A.Text text77 = new A.Text();
            text77.Text = "Second level";

            run61.Append(runProperties77);
            run61.Append(text77);

            paragraph62.Append(paragraphProperties35);
            paragraph62.Append(run61);

            A.Paragraph paragraph63 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties36 = new A.ParagraphProperties(){ Level = 2 };

            A.Run run62 = new A.Run();
            A.RunProperties runProperties78 = new A.RunProperties(){ Language = "en-GB" };
            A.Text text78 = new A.Text();
            text78.Text = "Third level";

            run62.Append(runProperties78);
            run62.Append(text78);

            paragraph63.Append(paragraphProperties36);
            paragraph63.Append(run62);

            A.Paragraph paragraph64 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties37 = new A.ParagraphProperties(){ Level = 3 };

            A.Run run63 = new A.Run();
            A.RunProperties runProperties79 = new A.RunProperties(){ Language = "en-GB" };
            A.Text text79 = new A.Text();
            text79.Text = "Fourth level";

            run63.Append(runProperties79);
            run63.Append(text79);

            paragraph64.Append(paragraphProperties37);
            paragraph64.Append(run63);

            A.Paragraph paragraph65 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties38 = new A.ParagraphProperties(){ Level = 4 };

            A.Run run64 = new A.Run();
            A.RunProperties runProperties80 = new A.RunProperties(){ Language = "en-GB" };
            A.Text text80 = new A.Text();
            text80.Text = "Fifth level";

            run64.Append(runProperties80);
            run64.Append(text80);
            A.EndParagraphRunProperties endParagraphRunProperties41 = new A.EndParagraphRunProperties(){ Language = "en-US", Dirty = false };

            paragraph65.Append(paragraphProperties38);
            paragraph65.Append(run64);
            paragraph65.Append(endParagraphRunProperties41);

            textBody45.Append(bodyProperties45);
            textBody45.Append(listStyle45);
            textBody45.Append(paragraph61);
            textBody45.Append(paragraph62);
            textBody45.Append(paragraph63);
            textBody45.Append(paragraph64);
            textBody45.Append(paragraph65);

            shape45.Append(nonVisualShapeProperties45);
            shape45.Append(shapeProperties46);
            shape45.Append(textBody45);

            Shape shape46 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties46 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties57 = new NonVisualDrawingProperties(){ Id = (UInt32Value)5U, Name = "Text Placeholder 4" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties46 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks45 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties46.Append(shapeLocks45);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties57 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape45 = new PlaceholderShape(){ Type = PlaceholderValues.Body, Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)3U };

            applicationNonVisualDrawingProperties57.Append(placeholderShape45);

            nonVisualShapeProperties46.Append(nonVisualDrawingProperties57);
            nonVisualShapeProperties46.Append(nonVisualShapeDrawingProperties46);
            nonVisualShapeProperties46.Append(applicationNonVisualDrawingProperties57);

            ShapeProperties shapeProperties47 = new ShapeProperties();

            A.Transform2D transform2D23 = new A.Transform2D();
            A.Offset offset33 = new A.Offset(){ X = 3857625L, Y = 1400969L };
            A.Extents extents33 = new A.Extents(){ Cx = 3239493L, Cy = 686593L };

            transform2D23.Append(offset33);
            transform2D23.Append(extents33);

            shapeProperties47.Append(transform2D23);

            TextBody textBody46 = new TextBody();
            A.BodyProperties bodyProperties46 = new A.BodyProperties(){ Anchor = A.TextAnchoringTypeValues.Bottom };

            A.ListStyle listStyle46 = new A.ListStyle();

            A.Level1ParagraphProperties level1ParagraphProperties18 = new A.Level1ParagraphProperties(){ LeftMargin = 0, Indent = 0 };
            A.NoBullet noBullet52 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties87 = new A.DefaultRunProperties(){ FontSize = 2000, Bold = true };

            level1ParagraphProperties18.Append(noBullet52);
            level1ParagraphProperties18.Append(defaultRunProperties87);

            A.Level2ParagraphProperties level2ParagraphProperties10 = new A.Level2ParagraphProperties(){ LeftMargin = 380985, Indent = 0 };
            A.NoBullet noBullet53 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties88 = new A.DefaultRunProperties(){ FontSize = 1667, Bold = true };

            level2ParagraphProperties10.Append(noBullet53);
            level2ParagraphProperties10.Append(defaultRunProperties88);

            A.Level3ParagraphProperties level3ParagraphProperties10 = new A.Level3ParagraphProperties(){ LeftMargin = 761970, Indent = 0 };
            A.NoBullet noBullet54 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties89 = new A.DefaultRunProperties(){ FontSize = 1500, Bold = true };

            level3ParagraphProperties10.Append(noBullet54);
            level3ParagraphProperties10.Append(defaultRunProperties89);

            A.Level4ParagraphProperties level4ParagraphProperties10 = new A.Level4ParagraphProperties(){ LeftMargin = 1142954, Indent = 0 };
            A.NoBullet noBullet55 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties90 = new A.DefaultRunProperties(){ FontSize = 1333, Bold = true };

            level4ParagraphProperties10.Append(noBullet55);
            level4ParagraphProperties10.Append(defaultRunProperties90);

            A.Level5ParagraphProperties level5ParagraphProperties10 = new A.Level5ParagraphProperties(){ LeftMargin = 1523939, Indent = 0 };
            A.NoBullet noBullet56 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties91 = new A.DefaultRunProperties(){ FontSize = 1333, Bold = true };

            level5ParagraphProperties10.Append(noBullet56);
            level5ParagraphProperties10.Append(defaultRunProperties91);

            A.Level6ParagraphProperties level6ParagraphProperties9 = new A.Level6ParagraphProperties(){ LeftMargin = 1904924, Indent = 0 };
            A.NoBullet noBullet57 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties92 = new A.DefaultRunProperties(){ FontSize = 1333, Bold = true };

            level6ParagraphProperties9.Append(noBullet57);
            level6ParagraphProperties9.Append(defaultRunProperties92);

            A.Level7ParagraphProperties level7ParagraphProperties9 = new A.Level7ParagraphProperties(){ LeftMargin = 2285909, Indent = 0 };
            A.NoBullet noBullet58 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties93 = new A.DefaultRunProperties(){ FontSize = 1333, Bold = true };

            level7ParagraphProperties9.Append(noBullet58);
            level7ParagraphProperties9.Append(defaultRunProperties93);

            A.Level8ParagraphProperties level8ParagraphProperties9 = new A.Level8ParagraphProperties(){ LeftMargin = 2666893, Indent = 0 };
            A.NoBullet noBullet59 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties94 = new A.DefaultRunProperties(){ FontSize = 1333, Bold = true };

            level8ParagraphProperties9.Append(noBullet59);
            level8ParagraphProperties9.Append(defaultRunProperties94);

            A.Level9ParagraphProperties level9ParagraphProperties9 = new A.Level9ParagraphProperties(){ LeftMargin = 3047878, Indent = 0 };
            A.NoBullet noBullet60 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties95 = new A.DefaultRunProperties(){ FontSize = 1333, Bold = true };

            level9ParagraphProperties9.Append(noBullet60);
            level9ParagraphProperties9.Append(defaultRunProperties95);

            listStyle46.Append(level1ParagraphProperties18);
            listStyle46.Append(level2ParagraphProperties10);
            listStyle46.Append(level3ParagraphProperties10);
            listStyle46.Append(level4ParagraphProperties10);
            listStyle46.Append(level5ParagraphProperties10);
            listStyle46.Append(level6ParagraphProperties9);
            listStyle46.Append(level7ParagraphProperties9);
            listStyle46.Append(level8ParagraphProperties9);
            listStyle46.Append(level9ParagraphProperties9);

            A.Paragraph paragraph66 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties39 = new A.ParagraphProperties(){ Level = 0 };

            A.Run run65 = new A.Run();
            A.RunProperties runProperties81 = new A.RunProperties(){ Language = "en-GB" };
            A.Text text81 = new A.Text();
            text81.Text = "Click to edit Master text styles";

            run65.Append(runProperties81);
            run65.Append(text81);

            paragraph66.Append(paragraphProperties39);
            paragraph66.Append(run65);

            textBody46.Append(bodyProperties46);
            textBody46.Append(listStyle46);
            textBody46.Append(paragraph66);

            shape46.Append(nonVisualShapeProperties46);
            shape46.Append(shapeProperties47);
            shape46.Append(textBody46);

            Shape shape47 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties47 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties58 = new NonVisualDrawingProperties(){ Id = (UInt32Value)6U, Name = "Content Placeholder 5" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties47 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks46 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties47.Append(shapeLocks46);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties58 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape46 = new PlaceholderShape(){ Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)4U };

            applicationNonVisualDrawingProperties58.Append(placeholderShape46);

            nonVisualShapeProperties47.Append(nonVisualDrawingProperties58);
            nonVisualShapeProperties47.Append(nonVisualShapeDrawingProperties47);
            nonVisualShapeProperties47.Append(applicationNonVisualDrawingProperties58);

            ShapeProperties shapeProperties48 = new ShapeProperties();

            A.Transform2D transform2D24 = new A.Transform2D();
            A.Offset offset34 = new A.Offset(){ X = 3857625L, Y = 2087563L };
            A.Extents extents34 = new A.Extents(){ Cx = 3239493L, Cy = 3070490L };

            transform2D24.Append(offset34);
            transform2D24.Append(extents34);

            shapeProperties48.Append(transform2D24);

            TextBody textBody47 = new TextBody();
            A.BodyProperties bodyProperties47 = new A.BodyProperties();
            A.ListStyle listStyle47 = new A.ListStyle();

            A.Paragraph paragraph67 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties40 = new A.ParagraphProperties(){ Level = 0 };

            A.Run run66 = new A.Run();
            A.RunProperties runProperties82 = new A.RunProperties(){ Language = "en-GB" };
            A.Text text82 = new A.Text();
            text82.Text = "Click to edit Master text styles";

            run66.Append(runProperties82);
            run66.Append(text82);

            paragraph67.Append(paragraphProperties40);
            paragraph67.Append(run66);

            A.Paragraph paragraph68 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties41 = new A.ParagraphProperties(){ Level = 1 };

            A.Run run67 = new A.Run();
            A.RunProperties runProperties83 = new A.RunProperties(){ Language = "en-GB" };
            A.Text text83 = new A.Text();
            text83.Text = "Second level";

            run67.Append(runProperties83);
            run67.Append(text83);

            paragraph68.Append(paragraphProperties41);
            paragraph68.Append(run67);

            A.Paragraph paragraph69 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties42 = new A.ParagraphProperties(){ Level = 2 };

            A.Run run68 = new A.Run();
            A.RunProperties runProperties84 = new A.RunProperties(){ Language = "en-GB" };
            A.Text text84 = new A.Text();
            text84.Text = "Third level";

            run68.Append(runProperties84);
            run68.Append(text84);

            paragraph69.Append(paragraphProperties42);
            paragraph69.Append(run68);

            A.Paragraph paragraph70 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties43 = new A.ParagraphProperties(){ Level = 3 };

            A.Run run69 = new A.Run();
            A.RunProperties runProperties85 = new A.RunProperties(){ Language = "en-GB" };
            A.Text text85 = new A.Text();
            text85.Text = "Fourth level";

            run69.Append(runProperties85);
            run69.Append(text85);

            paragraph70.Append(paragraphProperties43);
            paragraph70.Append(run69);

            A.Paragraph paragraph71 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties44 = new A.ParagraphProperties(){ Level = 4 };

            A.Run run70 = new A.Run();
            A.RunProperties runProperties86 = new A.RunProperties(){ Language = "en-GB" };
            A.Text text86 = new A.Text();
            text86.Text = "Fifth level";

            run70.Append(runProperties86);
            run70.Append(text86);
            A.EndParagraphRunProperties endParagraphRunProperties42 = new A.EndParagraphRunProperties(){ Language = "en-US", Dirty = false };

            paragraph71.Append(paragraphProperties44);
            paragraph71.Append(run70);
            paragraph71.Append(endParagraphRunProperties42);

            textBody47.Append(bodyProperties47);
            textBody47.Append(listStyle47);
            textBody47.Append(paragraph67);
            textBody47.Append(paragraph68);
            textBody47.Append(paragraph69);
            textBody47.Append(paragraph70);
            textBody47.Append(paragraph71);

            shape47.Append(nonVisualShapeProperties47);
            shape47.Append(shapeProperties48);
            shape47.Append(textBody47);

            Shape shape48 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties48 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties59 = new NonVisualDrawingProperties(){ Id = (UInt32Value)7U, Name = "Date Placeholder 6" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties48 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks47 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties48.Append(shapeLocks47);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties59 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape47 = new PlaceholderShape(){ Type = PlaceholderValues.DateAndTime, Size = PlaceholderSizeValues.Half, Index = (UInt32Value)10U };

            applicationNonVisualDrawingProperties59.Append(placeholderShape47);

            nonVisualShapeProperties48.Append(nonVisualDrawingProperties59);
            nonVisualShapeProperties48.Append(nonVisualShapeDrawingProperties48);
            nonVisualShapeProperties48.Append(applicationNonVisualDrawingProperties59);
            ShapeProperties shapeProperties49 = new ShapeProperties();

            TextBody textBody48 = new TextBody();
            A.BodyProperties bodyProperties48 = new A.BodyProperties();
            A.ListStyle listStyle48 = new A.ListStyle();

            A.Paragraph paragraph72 = new A.Paragraph();

            A.Field field17 = new A.Field(){ Id = "{A6EF8444-F3DD-4257-A539-5FF13F7A3AEA}", Type = "datetime1" };

            A.RunProperties runProperties87 = new A.RunProperties(){ Language = "en-US" };
            runProperties87.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            A.Text text87 = new A.Text();
            text87.Text = "1/17/2023";

            field17.Append(runProperties87);
            field17.Append(text87);
            A.EndParagraphRunProperties endParagraphRunProperties43 = new A.EndParagraphRunProperties(){ Language = "en-US", Dirty = false };

            paragraph72.Append(field17);
            paragraph72.Append(endParagraphRunProperties43);

            textBody48.Append(bodyProperties48);
            textBody48.Append(listStyle48);
            textBody48.Append(paragraph72);

            shape48.Append(nonVisualShapeProperties48);
            shape48.Append(shapeProperties49);
            shape48.Append(textBody48);

            Shape shape49 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties49 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties60 = new NonVisualDrawingProperties(){ Id = (UInt32Value)8U, Name = "Footer Placeholder 7" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties49 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks48 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties49.Append(shapeLocks48);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties60 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape48 = new PlaceholderShape(){ Type = PlaceholderValues.Footer, Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)11U };

            applicationNonVisualDrawingProperties60.Append(placeholderShape48);

            nonVisualShapeProperties49.Append(nonVisualDrawingProperties60);
            nonVisualShapeProperties49.Append(nonVisualShapeDrawingProperties49);
            nonVisualShapeProperties49.Append(applicationNonVisualDrawingProperties60);
            ShapeProperties shapeProperties50 = new ShapeProperties();

            TextBody textBody49 = new TextBody();
            A.BodyProperties bodyProperties49 = new A.BodyProperties();
            A.ListStyle listStyle49 = new A.ListStyle();

            A.Paragraph paragraph73 = new A.Paragraph();

            A.Run run71 = new A.Run();
            A.RunProperties runProperties88 = new A.RunProperties(){ Language = "en-US" };
            A.Text text88 = new A.Text();
            text88.Text = "Commercial & Workout Details";

            run71.Append(runProperties88);
            run71.Append(text88);
            A.EndParagraphRunProperties endParagraphRunProperties44 = new A.EndParagraphRunProperties(){ Language = "en-US", Dirty = false };

            paragraph73.Append(run71);
            paragraph73.Append(endParagraphRunProperties44);

            textBody49.Append(bodyProperties49);
            textBody49.Append(listStyle49);
            textBody49.Append(paragraph73);

            shape49.Append(nonVisualShapeProperties49);
            shape49.Append(shapeProperties50);
            shape49.Append(textBody49);

            Shape shape50 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties50 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties61 = new NonVisualDrawingProperties(){ Id = (UInt32Value)9U, Name = "Slide Number Placeholder 8" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties50 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks49 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties50.Append(shapeLocks49);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties61 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape49 = new PlaceholderShape(){ Type = PlaceholderValues.SlideNumber, Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)12U };

            applicationNonVisualDrawingProperties61.Append(placeholderShape49);

            nonVisualShapeProperties50.Append(nonVisualDrawingProperties61);
            nonVisualShapeProperties50.Append(nonVisualShapeDrawingProperties50);
            nonVisualShapeProperties50.Append(applicationNonVisualDrawingProperties61);
            ShapeProperties shapeProperties51 = new ShapeProperties();

            TextBody textBody50 = new TextBody();
            A.BodyProperties bodyProperties50 = new A.BodyProperties();
            A.ListStyle listStyle50 = new A.ListStyle();

            A.Paragraph paragraph74 = new A.Paragraph();

            A.Run run72 = new A.Run();

            A.RunProperties runProperties89 = new A.RunProperties(){ Language = "en-US" };

            A.SolidFill solidFill47 = new A.SolidFill();
            A.SchemeColor schemeColor63 = new A.SchemeColor(){ Val = A.SchemeColorValues.Accent2 };

            solidFill47.Append(schemeColor63);

            runProperties89.Append(solidFill47);
            A.Text text89 = new A.Text();
            text89.Text = "|";

            run72.Append(runProperties89);
            run72.Append(text89);

            A.Run run73 = new A.Run();
            A.RunProperties runProperties90 = new A.RunProperties(){ Language = "en-US" };
            A.Text text90 = new A.Text();
            text90.Text = "";

            run73.Append(runProperties90);
            run73.Append(text90);

            A.Field field18 = new A.Field(){ Id = "{E4F84C54-E2A4-46FF-B5B0-8F7A23C41D82}", Type = "slidenum" };

            A.RunProperties runProperties91 = new A.RunProperties(){ Language = "cs-CZ" };
            runProperties91.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            A.ParagraphProperties paragraphProperties45 = new A.ParagraphProperties();
            A.Text text91 = new A.Text();
            text91.Text = "‹#›";

            field18.Append(runProperties91);
            field18.Append(paragraphProperties45);
            field18.Append(text91);
            A.EndParagraphRunProperties endParagraphRunProperties45 = new A.EndParagraphRunProperties(){ Language = "cs-CZ", Dirty = false };

            paragraph74.Append(run72);
            paragraph74.Append(run73);
            paragraph74.Append(field18);
            paragraph74.Append(endParagraphRunProperties45);

            textBody50.Append(bodyProperties50);
            textBody50.Append(listStyle50);
            textBody50.Append(paragraph74);

            shape50.Append(nonVisualShapeProperties50);
            shape50.Append(shapeProperties51);
            shape50.Append(textBody50);

            shapeTree10.Append(nonVisualGroupShapeProperties10);
            shapeTree10.Append(groupShapeProperties10);
            shapeTree10.Append(shape43);
            shapeTree10.Append(shape44);
            shapeTree10.Append(shape45);
            shapeTree10.Append(shape46);
            shapeTree10.Append(shape47);
            shapeTree10.Append(shape48);
            shapeTree10.Append(shape49);
            shapeTree10.Append(shape50);

            CommonSlideDataExtensionList commonSlideDataExtensionList10 = new CommonSlideDataExtensionList();

            CommonSlideDataExtension commonSlideDataExtension10 = new CommonSlideDataExtension(){ Uri = "{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}" };

            P14.CreationId creationId10 = new P14.CreationId(){ Val = (UInt32Value)775669688U };
            creationId10.AddNamespaceDeclaration("p14", "http://schemas.microsoft.com/office/powerpoint/2010/main");

            commonSlideDataExtension10.Append(creationId10);

            commonSlideDataExtensionList10.Append(commonSlideDataExtension10);

            commonSlideData10.Append(shapeTree10);
            commonSlideData10.Append(commonSlideDataExtensionList10);

            ColorMapOverride colorMapOverride9 = new ColorMapOverride();
            A.MasterColorMapping masterColorMapping9 = new A.MasterColorMapping();

            colorMapOverride9.Append(masterColorMapping9);

            slideLayout9.Append(commonSlideData10);
            slideLayout9.Append(colorMapOverride9);

            slideLayoutPart9.SlideLayout = slideLayout9;
        }

        // Generates content of slideLayoutPart10.
        private void GenerateSlideLayoutPart10Content(SlideLayoutPart slideLayoutPart10)
        {
            SlideLayout slideLayout10 = new SlideLayout(){ Type = SlideLayoutValues.VerticalText, Preserve = true };
            slideLayout10.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            slideLayout10.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            slideLayout10.AddNamespaceDeclaration("p", "http://schemas.openxmlformats.org/presentationml/2006/main");

            CommonSlideData commonSlideData11 = new CommonSlideData(){ Name = "Title and Vertical Text" };

            ShapeTree shapeTree11 = new ShapeTree();

            NonVisualGroupShapeProperties nonVisualGroupShapeProperties11 = new NonVisualGroupShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties62 = new NonVisualDrawingProperties(){ Id = (UInt32Value)1U, Name = "" };
            NonVisualGroupShapeDrawingProperties nonVisualGroupShapeDrawingProperties11 = new NonVisualGroupShapeDrawingProperties();
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties62 = new ApplicationNonVisualDrawingProperties();

            nonVisualGroupShapeProperties11.Append(nonVisualDrawingProperties62);
            nonVisualGroupShapeProperties11.Append(nonVisualGroupShapeDrawingProperties11);
            nonVisualGroupShapeProperties11.Append(applicationNonVisualDrawingProperties62);

            GroupShapeProperties groupShapeProperties11 = new GroupShapeProperties();

            A.TransformGroup transformGroup11 = new A.TransformGroup();
            A.Offset offset35 = new A.Offset(){ X = 0L, Y = 0L };
            A.Extents extents35 = new A.Extents(){ Cx = 0L, Cy = 0L };
            A.ChildOffset childOffset11 = new A.ChildOffset(){ X = 0L, Y = 0L };
            A.ChildExtents childExtents11 = new A.ChildExtents(){ Cx = 0L, Cy = 0L };

            transformGroup11.Append(offset35);
            transformGroup11.Append(extents35);
            transformGroup11.Append(childOffset11);
            transformGroup11.Append(childExtents11);

            groupShapeProperties11.Append(transformGroup11);

            Shape shape51 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties51 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties63 = new NonVisualDrawingProperties(){ Id = (UInt32Value)2U, Name = "Title 1" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties51 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks50 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties51.Append(shapeLocks50);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties63 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape50 = new PlaceholderShape(){ Type = PlaceholderValues.Title };

            applicationNonVisualDrawingProperties63.Append(placeholderShape50);

            nonVisualShapeProperties51.Append(nonVisualDrawingProperties63);
            nonVisualShapeProperties51.Append(nonVisualShapeDrawingProperties51);
            nonVisualShapeProperties51.Append(applicationNonVisualDrawingProperties63);
            ShapeProperties shapeProperties52 = new ShapeProperties();

            TextBody textBody51 = new TextBody();
            A.BodyProperties bodyProperties51 = new A.BodyProperties();
            A.ListStyle listStyle51 = new A.ListStyle();

            A.Paragraph paragraph75 = new A.Paragraph();

            A.Run run74 = new A.Run();
            A.RunProperties runProperties92 = new A.RunProperties(){ Language = "en-GB" };
            A.Text text92 = new A.Text();
            text92.Text = "Click to edit Master title style";

            run74.Append(runProperties92);
            run74.Append(text92);
            A.EndParagraphRunProperties endParagraphRunProperties46 = new A.EndParagraphRunProperties(){ Language = "en-US", Dirty = false };

            paragraph75.Append(run74);
            paragraph75.Append(endParagraphRunProperties46);

            textBody51.Append(bodyProperties51);
            textBody51.Append(listStyle51);
            textBody51.Append(paragraph75);

            shape51.Append(nonVisualShapeProperties51);
            shape51.Append(shapeProperties52);
            shape51.Append(textBody51);

            Shape shape52 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties52 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties64 = new NonVisualDrawingProperties(){ Id = (UInt32Value)3U, Name = "Vertical Text Placeholder 2" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties52 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks51 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties52.Append(shapeLocks51);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties64 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape51 = new PlaceholderShape(){ Type = PlaceholderValues.Body, Orientation = DirectionValues.Vertical, Index = (UInt32Value)1U };

            applicationNonVisualDrawingProperties64.Append(placeholderShape51);

            nonVisualShapeProperties52.Append(nonVisualDrawingProperties64);
            nonVisualShapeProperties52.Append(nonVisualShapeDrawingProperties52);
            nonVisualShapeProperties52.Append(applicationNonVisualDrawingProperties64);
            ShapeProperties shapeProperties53 = new ShapeProperties();

            TextBody textBody52 = new TextBody();
            A.BodyProperties bodyProperties52 = new A.BodyProperties(){ Vertical = A.TextVerticalValues.EastAsianVetical };
            A.ListStyle listStyle52 = new A.ListStyle();

            A.Paragraph paragraph76 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties46 = new A.ParagraphProperties(){ Level = 0 };

            A.Run run75 = new A.Run();
            A.RunProperties runProperties93 = new A.RunProperties(){ Language = "en-GB" };
            A.Text text93 = new A.Text();
            text93.Text = "Click to edit Master text styles";

            run75.Append(runProperties93);
            run75.Append(text93);

            paragraph76.Append(paragraphProperties46);
            paragraph76.Append(run75);

            A.Paragraph paragraph77 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties47 = new A.ParagraphProperties(){ Level = 1 };

            A.Run run76 = new A.Run();
            A.RunProperties runProperties94 = new A.RunProperties(){ Language = "en-GB" };
            A.Text text94 = new A.Text();
            text94.Text = "Second level";

            run76.Append(runProperties94);
            run76.Append(text94);

            paragraph77.Append(paragraphProperties47);
            paragraph77.Append(run76);

            A.Paragraph paragraph78 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties48 = new A.ParagraphProperties(){ Level = 2 };

            A.Run run77 = new A.Run();
            A.RunProperties runProperties95 = new A.RunProperties(){ Language = "en-GB" };
            A.Text text95 = new A.Text();
            text95.Text = "Third level";

            run77.Append(runProperties95);
            run77.Append(text95);

            paragraph78.Append(paragraphProperties48);
            paragraph78.Append(run77);

            A.Paragraph paragraph79 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties49 = new A.ParagraphProperties(){ Level = 3 };

            A.Run run78 = new A.Run();
            A.RunProperties runProperties96 = new A.RunProperties(){ Language = "en-GB" };
            A.Text text96 = new A.Text();
            text96.Text = "Fourth level";

            run78.Append(runProperties96);
            run78.Append(text96);

            paragraph79.Append(paragraphProperties49);
            paragraph79.Append(run78);

            A.Paragraph paragraph80 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties50 = new A.ParagraphProperties(){ Level = 4 };

            A.Run run79 = new A.Run();
            A.RunProperties runProperties97 = new A.RunProperties(){ Language = "en-GB" };
            A.Text text97 = new A.Text();
            text97.Text = "Fifth level";

            run79.Append(runProperties97);
            run79.Append(text97);
            A.EndParagraphRunProperties endParagraphRunProperties47 = new A.EndParagraphRunProperties(){ Language = "en-US", Dirty = false };

            paragraph80.Append(paragraphProperties50);
            paragraph80.Append(run79);
            paragraph80.Append(endParagraphRunProperties47);

            textBody52.Append(bodyProperties52);
            textBody52.Append(listStyle52);
            textBody52.Append(paragraph76);
            textBody52.Append(paragraph77);
            textBody52.Append(paragraph78);
            textBody52.Append(paragraph79);
            textBody52.Append(paragraph80);

            shape52.Append(nonVisualShapeProperties52);
            shape52.Append(shapeProperties53);
            shape52.Append(textBody52);

            Shape shape53 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties53 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties65 = new NonVisualDrawingProperties(){ Id = (UInt32Value)4U, Name = "Date Placeholder 3" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties53 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks52 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties53.Append(shapeLocks52);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties65 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape52 = new PlaceholderShape(){ Type = PlaceholderValues.DateAndTime, Size = PlaceholderSizeValues.Half, Index = (UInt32Value)10U };

            applicationNonVisualDrawingProperties65.Append(placeholderShape52);

            nonVisualShapeProperties53.Append(nonVisualDrawingProperties65);
            nonVisualShapeProperties53.Append(nonVisualShapeDrawingProperties53);
            nonVisualShapeProperties53.Append(applicationNonVisualDrawingProperties65);
            ShapeProperties shapeProperties54 = new ShapeProperties();

            TextBody textBody53 = new TextBody();
            A.BodyProperties bodyProperties53 = new A.BodyProperties();
            A.ListStyle listStyle53 = new A.ListStyle();

            A.Paragraph paragraph81 = new A.Paragraph();

            A.Field field19 = new A.Field(){ Id = "{3156E602-F80F-45BA-BBA5-374C402D5ACF}", Type = "datetime1" };

            A.RunProperties runProperties98 = new A.RunProperties(){ Language = "en-US" };
            runProperties98.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            A.Text text98 = new A.Text();
            text98.Text = "1/17/2023";

            field19.Append(runProperties98);
            field19.Append(text98);
            A.EndParagraphRunProperties endParagraphRunProperties48 = new A.EndParagraphRunProperties(){ Language = "en-US", Dirty = false };

            paragraph81.Append(field19);
            paragraph81.Append(endParagraphRunProperties48);

            textBody53.Append(bodyProperties53);
            textBody53.Append(listStyle53);
            textBody53.Append(paragraph81);

            shape53.Append(nonVisualShapeProperties53);
            shape53.Append(shapeProperties54);
            shape53.Append(textBody53);

            Shape shape54 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties54 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties66 = new NonVisualDrawingProperties(){ Id = (UInt32Value)5U, Name = "Footer Placeholder 4" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties54 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks53 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties54.Append(shapeLocks53);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties66 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape53 = new PlaceholderShape(){ Type = PlaceholderValues.Footer, Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)11U };

            applicationNonVisualDrawingProperties66.Append(placeholderShape53);

            nonVisualShapeProperties54.Append(nonVisualDrawingProperties66);
            nonVisualShapeProperties54.Append(nonVisualShapeDrawingProperties54);
            nonVisualShapeProperties54.Append(applicationNonVisualDrawingProperties66);
            ShapeProperties shapeProperties55 = new ShapeProperties();

            TextBody textBody54 = new TextBody();
            A.BodyProperties bodyProperties54 = new A.BodyProperties();
            A.ListStyle listStyle54 = new A.ListStyle();

            A.Paragraph paragraph82 = new A.Paragraph();

            A.Run run80 = new A.Run();
            A.RunProperties runProperties99 = new A.RunProperties(){ Language = "en-US" };
            A.Text text99 = new A.Text();
            text99.Text = "Commercial & Workout Details";

            run80.Append(runProperties99);
            run80.Append(text99);
            A.EndParagraphRunProperties endParagraphRunProperties49 = new A.EndParagraphRunProperties(){ Language = "en-US", Dirty = false };

            paragraph82.Append(run80);
            paragraph82.Append(endParagraphRunProperties49);

            textBody54.Append(bodyProperties54);
            textBody54.Append(listStyle54);
            textBody54.Append(paragraph82);

            shape54.Append(nonVisualShapeProperties54);
            shape54.Append(shapeProperties55);
            shape54.Append(textBody54);

            Shape shape55 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties55 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties67 = new NonVisualDrawingProperties(){ Id = (UInt32Value)6U, Name = "Slide Number Placeholder 5" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties55 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks54 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties55.Append(shapeLocks54);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties67 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape54 = new PlaceholderShape(){ Type = PlaceholderValues.SlideNumber, Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)12U };

            applicationNonVisualDrawingProperties67.Append(placeholderShape54);

            nonVisualShapeProperties55.Append(nonVisualDrawingProperties67);
            nonVisualShapeProperties55.Append(nonVisualShapeDrawingProperties55);
            nonVisualShapeProperties55.Append(applicationNonVisualDrawingProperties67);
            ShapeProperties shapeProperties56 = new ShapeProperties();

            TextBody textBody55 = new TextBody();
            A.BodyProperties bodyProperties55 = new A.BodyProperties();
            A.ListStyle listStyle55 = new A.ListStyle();

            A.Paragraph paragraph83 = new A.Paragraph();

            A.Run run81 = new A.Run();

            A.RunProperties runProperties100 = new A.RunProperties(){ Language = "en-US" };

            A.SolidFill solidFill48 = new A.SolidFill();
            A.SchemeColor schemeColor64 = new A.SchemeColor(){ Val = A.SchemeColorValues.Accent2 };

            solidFill48.Append(schemeColor64);

            runProperties100.Append(solidFill48);
            A.Text text100 = new A.Text();
            text100.Text = "|";

            run81.Append(runProperties100);
            run81.Append(text100);

            A.Run run82 = new A.Run();
            A.RunProperties runProperties101 = new A.RunProperties(){ Language = "en-US" };
            A.Text text101 = new A.Text();
            text101.Text = "";

            run82.Append(runProperties101);
            run82.Append(text101);

            A.Field field20 = new A.Field(){ Id = "{E4F84C54-E2A4-46FF-B5B0-8F7A23C41D82}", Type = "slidenum" };

            A.RunProperties runProperties102 = new A.RunProperties(){ Language = "cs-CZ" };
            runProperties102.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            A.ParagraphProperties paragraphProperties51 = new A.ParagraphProperties();
            A.Text text102 = new A.Text();
            text102.Text = "‹#›";

            field20.Append(runProperties102);
            field20.Append(paragraphProperties51);
            field20.Append(text102);
            A.EndParagraphRunProperties endParagraphRunProperties50 = new A.EndParagraphRunProperties(){ Language = "cs-CZ", Dirty = false };

            paragraph83.Append(run81);
            paragraph83.Append(run82);
            paragraph83.Append(field20);
            paragraph83.Append(endParagraphRunProperties50);

            textBody55.Append(bodyProperties55);
            textBody55.Append(listStyle55);
            textBody55.Append(paragraph83);

            shape55.Append(nonVisualShapeProperties55);
            shape55.Append(shapeProperties56);
            shape55.Append(textBody55);

            shapeTree11.Append(nonVisualGroupShapeProperties11);
            shapeTree11.Append(groupShapeProperties11);
            shapeTree11.Append(shape51);
            shapeTree11.Append(shape52);
            shapeTree11.Append(shape53);
            shapeTree11.Append(shape54);
            shapeTree11.Append(shape55);

            CommonSlideDataExtensionList commonSlideDataExtensionList11 = new CommonSlideDataExtensionList();

            CommonSlideDataExtension commonSlideDataExtension11 = new CommonSlideDataExtension(){ Uri = "{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}" };

            P14.CreationId creationId11 = new P14.CreationId(){ Val = (UInt32Value)4124392494U };
            creationId11.AddNamespaceDeclaration("p14", "http://schemas.microsoft.com/office/powerpoint/2010/main");

            commonSlideDataExtension11.Append(creationId11);

            commonSlideDataExtensionList11.Append(commonSlideDataExtension11);

            commonSlideData11.Append(shapeTree11);
            commonSlideData11.Append(commonSlideDataExtensionList11);

            ColorMapOverride colorMapOverride10 = new ColorMapOverride();
            A.MasterColorMapping masterColorMapping10 = new A.MasterColorMapping();

            colorMapOverride10.Append(masterColorMapping10);

            slideLayout10.Append(commonSlideData11);
            slideLayout10.Append(colorMapOverride10);

            slideLayoutPart10.SlideLayout = slideLayout10;
        }

        // Generates content of slideLayoutPart11.
        private void GenerateSlideLayoutPart11Content(SlideLayoutPart slideLayoutPart11)
        {
            SlideLayout slideLayout11 = new SlideLayout(){ Type = SlideLayoutValues.TwoObjects, Preserve = true };
            slideLayout11.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            slideLayout11.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            slideLayout11.AddNamespaceDeclaration("p", "http://schemas.openxmlformats.org/presentationml/2006/main");

            CommonSlideData commonSlideData12 = new CommonSlideData(){ Name = "Two Content" };

            ShapeTree shapeTree12 = new ShapeTree();

            NonVisualGroupShapeProperties nonVisualGroupShapeProperties12 = new NonVisualGroupShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties68 = new NonVisualDrawingProperties(){ Id = (UInt32Value)1U, Name = "" };
            NonVisualGroupShapeDrawingProperties nonVisualGroupShapeDrawingProperties12 = new NonVisualGroupShapeDrawingProperties();
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties68 = new ApplicationNonVisualDrawingProperties();

            nonVisualGroupShapeProperties12.Append(nonVisualDrawingProperties68);
            nonVisualGroupShapeProperties12.Append(nonVisualGroupShapeDrawingProperties12);
            nonVisualGroupShapeProperties12.Append(applicationNonVisualDrawingProperties68);

            GroupShapeProperties groupShapeProperties12 = new GroupShapeProperties();

            A.TransformGroup transformGroup12 = new A.TransformGroup();
            A.Offset offset36 = new A.Offset(){ X = 0L, Y = 0L };
            A.Extents extents36 = new A.Extents(){ Cx = 0L, Cy = 0L };
            A.ChildOffset childOffset12 = new A.ChildOffset(){ X = 0L, Y = 0L };
            A.ChildExtents childExtents12 = new A.ChildExtents(){ Cx = 0L, Cy = 0L };

            transformGroup12.Append(offset36);
            transformGroup12.Append(extents36);
            transformGroup12.Append(childOffset12);
            transformGroup12.Append(childExtents12);

            groupShapeProperties12.Append(transformGroup12);

            Shape shape56 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties56 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties69 = new NonVisualDrawingProperties(){ Id = (UInt32Value)2U, Name = "Title 1" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties56 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks55 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties56.Append(shapeLocks55);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties69 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape55 = new PlaceholderShape(){ Type = PlaceholderValues.Title };

            applicationNonVisualDrawingProperties69.Append(placeholderShape55);

            nonVisualShapeProperties56.Append(nonVisualDrawingProperties69);
            nonVisualShapeProperties56.Append(nonVisualShapeDrawingProperties56);
            nonVisualShapeProperties56.Append(applicationNonVisualDrawingProperties69);
            ShapeProperties shapeProperties57 = new ShapeProperties();

            TextBody textBody56 = new TextBody();
            A.BodyProperties bodyProperties56 = new A.BodyProperties();
            A.ListStyle listStyle56 = new A.ListStyle();

            A.Paragraph paragraph84 = new A.Paragraph();

            A.Run run83 = new A.Run();
            A.RunProperties runProperties103 = new A.RunProperties(){ Language = "en-GB" };
            A.Text text103 = new A.Text();
            text103.Text = "Click to edit Master title style";

            run83.Append(runProperties103);
            run83.Append(text103);
            A.EndParagraphRunProperties endParagraphRunProperties51 = new A.EndParagraphRunProperties(){ Language = "en-US", Dirty = false };

            paragraph84.Append(run83);
            paragraph84.Append(endParagraphRunProperties51);

            textBody56.Append(bodyProperties56);
            textBody56.Append(listStyle56);
            textBody56.Append(paragraph84);

            shape56.Append(nonVisualShapeProperties56);
            shape56.Append(shapeProperties57);
            shape56.Append(textBody56);

            Shape shape57 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties57 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties70 = new NonVisualDrawingProperties(){ Id = (UInt32Value)3U, Name = "Content Placeholder 2" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties57 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks56 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties57.Append(shapeLocks56);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties70 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape56 = new PlaceholderShape(){ Size = PlaceholderSizeValues.Half, Index = (UInt32Value)1U };

            applicationNonVisualDrawingProperties70.Append(placeholderShape56);

            nonVisualShapeProperties57.Append(nonVisualDrawingProperties70);
            nonVisualShapeProperties57.Append(nonVisualShapeDrawingProperties57);
            nonVisualShapeProperties57.Append(applicationNonVisualDrawingProperties70);

            ShapeProperties shapeProperties58 = new ShapeProperties();

            A.Transform2D transform2D25 = new A.Transform2D();
            A.Offset offset37 = new A.Offset(){ X = 523875L, Y = 1521354L };
            A.Extents extents37 = new A.Extents(){ Cx = 3238500L, Cy = 3626115L };

            transform2D25.Append(offset37);
            transform2D25.Append(extents37);

            shapeProperties58.Append(transform2D25);

            TextBody textBody57 = new TextBody();
            A.BodyProperties bodyProperties57 = new A.BodyProperties();
            A.ListStyle listStyle57 = new A.ListStyle();

            A.Paragraph paragraph85 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties52 = new A.ParagraphProperties(){ Level = 0 };

            A.Run run84 = new A.Run();
            A.RunProperties runProperties104 = new A.RunProperties(){ Language = "en-GB" };
            A.Text text104 = new A.Text();
            text104.Text = "Click to edit Master text styles";

            run84.Append(runProperties104);
            run84.Append(text104);

            paragraph85.Append(paragraphProperties52);
            paragraph85.Append(run84);

            A.Paragraph paragraph86 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties53 = new A.ParagraphProperties(){ Level = 1 };

            A.Run run85 = new A.Run();
            A.RunProperties runProperties105 = new A.RunProperties(){ Language = "en-GB" };
            A.Text text105 = new A.Text();
            text105.Text = "Second level";

            run85.Append(runProperties105);
            run85.Append(text105);

            paragraph86.Append(paragraphProperties53);
            paragraph86.Append(run85);

            A.Paragraph paragraph87 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties54 = new A.ParagraphProperties(){ Level = 2 };

            A.Run run86 = new A.Run();
            A.RunProperties runProperties106 = new A.RunProperties(){ Language = "en-GB" };
            A.Text text106 = new A.Text();
            text106.Text = "Third level";

            run86.Append(runProperties106);
            run86.Append(text106);

            paragraph87.Append(paragraphProperties54);
            paragraph87.Append(run86);

            A.Paragraph paragraph88 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties55 = new A.ParagraphProperties(){ Level = 3 };

            A.Run run87 = new A.Run();
            A.RunProperties runProperties107 = new A.RunProperties(){ Language = "en-GB" };
            A.Text text107 = new A.Text();
            text107.Text = "Fourth level";

            run87.Append(runProperties107);
            run87.Append(text107);

            paragraph88.Append(paragraphProperties55);
            paragraph88.Append(run87);

            A.Paragraph paragraph89 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties56 = new A.ParagraphProperties(){ Level = 4 };

            A.Run run88 = new A.Run();
            A.RunProperties runProperties108 = new A.RunProperties(){ Language = "en-GB" };
            A.Text text108 = new A.Text();
            text108.Text = "Fifth level";

            run88.Append(runProperties108);
            run88.Append(text108);
            A.EndParagraphRunProperties endParagraphRunProperties52 = new A.EndParagraphRunProperties(){ Language = "en-US", Dirty = false };

            paragraph89.Append(paragraphProperties56);
            paragraph89.Append(run88);
            paragraph89.Append(endParagraphRunProperties52);

            textBody57.Append(bodyProperties57);
            textBody57.Append(listStyle57);
            textBody57.Append(paragraph85);
            textBody57.Append(paragraph86);
            textBody57.Append(paragraph87);
            textBody57.Append(paragraph88);
            textBody57.Append(paragraph89);

            shape57.Append(nonVisualShapeProperties57);
            shape57.Append(shapeProperties58);
            shape57.Append(textBody57);

            Shape shape58 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties58 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties71 = new NonVisualDrawingProperties(){ Id = (UInt32Value)4U, Name = "Content Placeholder 3" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties58 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks57 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties58.Append(shapeLocks57);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties71 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape57 = new PlaceholderShape(){ Size = PlaceholderSizeValues.Half, Index = (UInt32Value)2U };

            applicationNonVisualDrawingProperties71.Append(placeholderShape57);

            nonVisualShapeProperties58.Append(nonVisualDrawingProperties71);
            nonVisualShapeProperties58.Append(nonVisualShapeDrawingProperties58);
            nonVisualShapeProperties58.Append(applicationNonVisualDrawingProperties71);

            ShapeProperties shapeProperties59 = new ShapeProperties();

            A.Transform2D transform2D26 = new A.Transform2D();
            A.Offset offset38 = new A.Offset(){ X = 3857625L, Y = 1521354L };
            A.Extents extents38 = new A.Extents(){ Cx = 3238500L, Cy = 3626115L };

            transform2D26.Append(offset38);
            transform2D26.Append(extents38);

            shapeProperties59.Append(transform2D26);

            TextBody textBody58 = new TextBody();
            A.BodyProperties bodyProperties58 = new A.BodyProperties();
            A.ListStyle listStyle58 = new A.ListStyle();

            A.Paragraph paragraph90 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties57 = new A.ParagraphProperties(){ Level = 0 };

            A.Run run89 = new A.Run();
            A.RunProperties runProperties109 = new A.RunProperties(){ Language = "en-GB" };
            A.Text text109 = new A.Text();
            text109.Text = "Click to edit Master text styles";

            run89.Append(runProperties109);
            run89.Append(text109);

            paragraph90.Append(paragraphProperties57);
            paragraph90.Append(run89);

            A.Paragraph paragraph91 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties58 = new A.ParagraphProperties(){ Level = 1 };

            A.Run run90 = new A.Run();
            A.RunProperties runProperties110 = new A.RunProperties(){ Language = "en-GB" };
            A.Text text110 = new A.Text();
            text110.Text = "Second level";

            run90.Append(runProperties110);
            run90.Append(text110);

            paragraph91.Append(paragraphProperties58);
            paragraph91.Append(run90);

            A.Paragraph paragraph92 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties59 = new A.ParagraphProperties(){ Level = 2 };

            A.Run run91 = new A.Run();
            A.RunProperties runProperties111 = new A.RunProperties(){ Language = "en-GB" };
            A.Text text111 = new A.Text();
            text111.Text = "Third level";

            run91.Append(runProperties111);
            run91.Append(text111);

            paragraph92.Append(paragraphProperties59);
            paragraph92.Append(run91);

            A.Paragraph paragraph93 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties60 = new A.ParagraphProperties(){ Level = 3 };

            A.Run run92 = new A.Run();
            A.RunProperties runProperties112 = new A.RunProperties(){ Language = "en-GB" };
            A.Text text112 = new A.Text();
            text112.Text = "Fourth level";

            run92.Append(runProperties112);
            run92.Append(text112);

            paragraph93.Append(paragraphProperties60);
            paragraph93.Append(run92);

            A.Paragraph paragraph94 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties61 = new A.ParagraphProperties(){ Level = 4 };

            A.Run run93 = new A.Run();
            A.RunProperties runProperties113 = new A.RunProperties(){ Language = "en-GB" };
            A.Text text113 = new A.Text();
            text113.Text = "Fifth level";

            run93.Append(runProperties113);
            run93.Append(text113);
            A.EndParagraphRunProperties endParagraphRunProperties53 = new A.EndParagraphRunProperties(){ Language = "en-US", Dirty = false };

            paragraph94.Append(paragraphProperties61);
            paragraph94.Append(run93);
            paragraph94.Append(endParagraphRunProperties53);

            textBody58.Append(bodyProperties58);
            textBody58.Append(listStyle58);
            textBody58.Append(paragraph90);
            textBody58.Append(paragraph91);
            textBody58.Append(paragraph92);
            textBody58.Append(paragraph93);
            textBody58.Append(paragraph94);

            shape58.Append(nonVisualShapeProperties58);
            shape58.Append(shapeProperties59);
            shape58.Append(textBody58);

            Shape shape59 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties59 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties72 = new NonVisualDrawingProperties(){ Id = (UInt32Value)5U, Name = "Date Placeholder 4" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties59 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks58 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties59.Append(shapeLocks58);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties72 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape58 = new PlaceholderShape(){ Type = PlaceholderValues.DateAndTime, Size = PlaceholderSizeValues.Half, Index = (UInt32Value)10U };

            applicationNonVisualDrawingProperties72.Append(placeholderShape58);

            nonVisualShapeProperties59.Append(nonVisualDrawingProperties72);
            nonVisualShapeProperties59.Append(nonVisualShapeDrawingProperties59);
            nonVisualShapeProperties59.Append(applicationNonVisualDrawingProperties72);
            ShapeProperties shapeProperties60 = new ShapeProperties();

            TextBody textBody59 = new TextBody();
            A.BodyProperties bodyProperties59 = new A.BodyProperties();
            A.ListStyle listStyle59 = new A.ListStyle();

            A.Paragraph paragraph95 = new A.Paragraph();

            A.Field field21 = new A.Field(){ Id = "{B904F67C-45CA-4694-AF10-9217F14C1969}", Type = "datetime1" };

            A.RunProperties runProperties114 = new A.RunProperties(){ Language = "en-US" };
            runProperties114.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            A.Text text114 = new A.Text();
            text114.Text = "1/17/2023";

            field21.Append(runProperties114);
            field21.Append(text114);
            A.EndParagraphRunProperties endParagraphRunProperties54 = new A.EndParagraphRunProperties(){ Language = "en-US", Dirty = false };

            paragraph95.Append(field21);
            paragraph95.Append(endParagraphRunProperties54);

            textBody59.Append(bodyProperties59);
            textBody59.Append(listStyle59);
            textBody59.Append(paragraph95);

            shape59.Append(nonVisualShapeProperties59);
            shape59.Append(shapeProperties60);
            shape59.Append(textBody59);

            Shape shape60 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties60 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties73 = new NonVisualDrawingProperties(){ Id = (UInt32Value)6U, Name = "Footer Placeholder 5" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties60 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks59 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties60.Append(shapeLocks59);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties73 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape59 = new PlaceholderShape(){ Type = PlaceholderValues.Footer, Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)11U };

            applicationNonVisualDrawingProperties73.Append(placeholderShape59);

            nonVisualShapeProperties60.Append(nonVisualDrawingProperties73);
            nonVisualShapeProperties60.Append(nonVisualShapeDrawingProperties60);
            nonVisualShapeProperties60.Append(applicationNonVisualDrawingProperties73);
            ShapeProperties shapeProperties61 = new ShapeProperties();

            TextBody textBody60 = new TextBody();
            A.BodyProperties bodyProperties60 = new A.BodyProperties();
            A.ListStyle listStyle60 = new A.ListStyle();

            A.Paragraph paragraph96 = new A.Paragraph();

            A.Run run94 = new A.Run();
            A.RunProperties runProperties115 = new A.RunProperties(){ Language = "en-US" };
            A.Text text115 = new A.Text();
            text115.Text = "Commercial & Workout Details";

            run94.Append(runProperties115);
            run94.Append(text115);
            A.EndParagraphRunProperties endParagraphRunProperties55 = new A.EndParagraphRunProperties(){ Language = "en-US", Dirty = false };

            paragraph96.Append(run94);
            paragraph96.Append(endParagraphRunProperties55);

            textBody60.Append(bodyProperties60);
            textBody60.Append(listStyle60);
            textBody60.Append(paragraph96);

            shape60.Append(nonVisualShapeProperties60);
            shape60.Append(shapeProperties61);
            shape60.Append(textBody60);

            Shape shape61 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties61 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties74 = new NonVisualDrawingProperties(){ Id = (UInt32Value)7U, Name = "Slide Number Placeholder 6" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties61 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks60 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties61.Append(shapeLocks60);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties74 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape60 = new PlaceholderShape(){ Type = PlaceholderValues.SlideNumber, Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)12U };

            applicationNonVisualDrawingProperties74.Append(placeholderShape60);

            nonVisualShapeProperties61.Append(nonVisualDrawingProperties74);
            nonVisualShapeProperties61.Append(nonVisualShapeDrawingProperties61);
            nonVisualShapeProperties61.Append(applicationNonVisualDrawingProperties74);
            ShapeProperties shapeProperties62 = new ShapeProperties();

            TextBody textBody61 = new TextBody();
            A.BodyProperties bodyProperties61 = new A.BodyProperties();
            A.ListStyle listStyle61 = new A.ListStyle();

            A.Paragraph paragraph97 = new A.Paragraph();

            A.Run run95 = new A.Run();

            A.RunProperties runProperties116 = new A.RunProperties(){ Language = "en-US" };

            A.SolidFill solidFill49 = new A.SolidFill();
            A.SchemeColor schemeColor65 = new A.SchemeColor(){ Val = A.SchemeColorValues.Accent2 };

            solidFill49.Append(schemeColor65);

            runProperties116.Append(solidFill49);
            A.Text text116 = new A.Text();
            text116.Text = "|";

            run95.Append(runProperties116);
            run95.Append(text116);

            A.Run run96 = new A.Run();
            A.RunProperties runProperties117 = new A.RunProperties(){ Language = "en-US" };
            A.Text text117 = new A.Text();
            text117.Text = "";

            run96.Append(runProperties117);
            run96.Append(text117);

            A.Field field22 = new A.Field(){ Id = "{E4F84C54-E2A4-46FF-B5B0-8F7A23C41D82}", Type = "slidenum" };

            A.RunProperties runProperties118 = new A.RunProperties(){ Language = "cs-CZ" };
            runProperties118.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            A.ParagraphProperties paragraphProperties62 = new A.ParagraphProperties();
            A.Text text118 = new A.Text();
            text118.Text = "‹#›";

            field22.Append(runProperties118);
            field22.Append(paragraphProperties62);
            field22.Append(text118);
            A.EndParagraphRunProperties endParagraphRunProperties56 = new A.EndParagraphRunProperties(){ Language = "cs-CZ", Dirty = false };

            paragraph97.Append(run95);
            paragraph97.Append(run96);
            paragraph97.Append(field22);
            paragraph97.Append(endParagraphRunProperties56);

            textBody61.Append(bodyProperties61);
            textBody61.Append(listStyle61);
            textBody61.Append(paragraph97);

            shape61.Append(nonVisualShapeProperties61);
            shape61.Append(shapeProperties62);
            shape61.Append(textBody61);

            shapeTree12.Append(nonVisualGroupShapeProperties12);
            shapeTree12.Append(groupShapeProperties12);
            shapeTree12.Append(shape56);
            shapeTree12.Append(shape57);
            shapeTree12.Append(shape58);
            shapeTree12.Append(shape59);
            shapeTree12.Append(shape60);
            shapeTree12.Append(shape61);

            CommonSlideDataExtensionList commonSlideDataExtensionList12 = new CommonSlideDataExtensionList();

            CommonSlideDataExtension commonSlideDataExtension12 = new CommonSlideDataExtension(){ Uri = "{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}" };

            P14.CreationId creationId12 = new P14.CreationId(){ Val = (UInt32Value)3124828198U };
            creationId12.AddNamespaceDeclaration("p14", "http://schemas.microsoft.com/office/powerpoint/2010/main");

            commonSlideDataExtension12.Append(creationId12);

            commonSlideDataExtensionList12.Append(commonSlideDataExtension12);

            commonSlideData12.Append(shapeTree12);
            commonSlideData12.Append(commonSlideDataExtensionList12);

            ColorMapOverride colorMapOverride11 = new ColorMapOverride();
            A.MasterColorMapping masterColorMapping11 = new A.MasterColorMapping();

            colorMapOverride11.Append(masterColorMapping11);

            slideLayout11.Append(commonSlideData12);
            slideLayout11.Append(colorMapOverride11);

            slideLayoutPart11.SlideLayout = slideLayout11;
        }

        // Generates content of slideLayoutPart12.
        private void GenerateSlideLayoutPart12Content(SlideLayoutPart slideLayoutPart12)
        {
            SlideLayout slideLayout12 = new SlideLayout(){ Type = SlideLayoutValues.PictureText, Preserve = true };
            slideLayout12.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            slideLayout12.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            slideLayout12.AddNamespaceDeclaration("p", "http://schemas.openxmlformats.org/presentationml/2006/main");

            CommonSlideData commonSlideData13 = new CommonSlideData(){ Name = "Picture with Caption" };

            ShapeTree shapeTree13 = new ShapeTree();

            NonVisualGroupShapeProperties nonVisualGroupShapeProperties13 = new NonVisualGroupShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties75 = new NonVisualDrawingProperties(){ Id = (UInt32Value)1U, Name = "" };
            NonVisualGroupShapeDrawingProperties nonVisualGroupShapeDrawingProperties13 = new NonVisualGroupShapeDrawingProperties();
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties75 = new ApplicationNonVisualDrawingProperties();

            nonVisualGroupShapeProperties13.Append(nonVisualDrawingProperties75);
            nonVisualGroupShapeProperties13.Append(nonVisualGroupShapeDrawingProperties13);
            nonVisualGroupShapeProperties13.Append(applicationNonVisualDrawingProperties75);

            GroupShapeProperties groupShapeProperties13 = new GroupShapeProperties();

            A.TransformGroup transformGroup13 = new A.TransformGroup();
            A.Offset offset39 = new A.Offset(){ X = 0L, Y = 0L };
            A.Extents extents39 = new A.Extents(){ Cx = 0L, Cy = 0L };
            A.ChildOffset childOffset13 = new A.ChildOffset(){ X = 0L, Y = 0L };
            A.ChildExtents childExtents13 = new A.ChildExtents(){ Cx = 0L, Cy = 0L };

            transformGroup13.Append(offset39);
            transformGroup13.Append(extents39);
            transformGroup13.Append(childOffset13);
            transformGroup13.Append(childExtents13);

            groupShapeProperties13.Append(transformGroup13);

            Shape shape62 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties62 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties76 = new NonVisualDrawingProperties(){ Id = (UInt32Value)2U, Name = "Title 1" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties62 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks61 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties62.Append(shapeLocks61);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties76 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape61 = new PlaceholderShape(){ Type = PlaceholderValues.Title };

            applicationNonVisualDrawingProperties76.Append(placeholderShape61);

            nonVisualShapeProperties62.Append(nonVisualDrawingProperties76);
            nonVisualShapeProperties62.Append(nonVisualShapeDrawingProperties62);
            nonVisualShapeProperties62.Append(applicationNonVisualDrawingProperties76);

            ShapeProperties shapeProperties63 = new ShapeProperties();

            A.Transform2D transform2D27 = new A.Transform2D();
            A.Offset offset40 = new A.Offset(){ X = 524868L, Y = 381000L };
            A.Extents extents40 = new A.Extents(){ Cx = 2457648L, Cy = 1333500L };

            transform2D27.Append(offset40);
            transform2D27.Append(extents40);

            shapeProperties63.Append(transform2D27);

            TextBody textBody62 = new TextBody();
            A.BodyProperties bodyProperties62 = new A.BodyProperties(){ Anchor = A.TextAnchoringTypeValues.Bottom };

            A.ListStyle listStyle62 = new A.ListStyle();

            A.Level1ParagraphProperties level1ParagraphProperties19 = new A.Level1ParagraphProperties();
            A.DefaultRunProperties defaultRunProperties96 = new A.DefaultRunProperties(){ FontSize = 2667 };

            level1ParagraphProperties19.Append(defaultRunProperties96);

            listStyle62.Append(level1ParagraphProperties19);

            A.Paragraph paragraph98 = new A.Paragraph();

            A.Run run97 = new A.Run();
            A.RunProperties runProperties119 = new A.RunProperties(){ Language = "en-GB" };
            A.Text text119 = new A.Text();
            text119.Text = "Click to edit Master title style";

            run97.Append(runProperties119);
            run97.Append(text119);
            A.EndParagraphRunProperties endParagraphRunProperties57 = new A.EndParagraphRunProperties(){ Language = "en-US", Dirty = false };

            paragraph98.Append(run97);
            paragraph98.Append(endParagraphRunProperties57);

            textBody62.Append(bodyProperties62);
            textBody62.Append(listStyle62);
            textBody62.Append(paragraph98);

            shape62.Append(nonVisualShapeProperties62);
            shape62.Append(shapeProperties63);
            shape62.Append(textBody62);

            Shape shape63 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties63 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties77 = new NonVisualDrawingProperties(){ Id = (UInt32Value)3U, Name = "Picture Placeholder 2" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties63 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks62 = new A.ShapeLocks(){ NoGrouping = true, NoChangeAspect = true };

            nonVisualShapeDrawingProperties63.Append(shapeLocks62);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties77 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape62 = new PlaceholderShape(){ Type = PlaceholderValues.Picture, Index = (UInt32Value)1U };

            applicationNonVisualDrawingProperties77.Append(placeholderShape62);

            nonVisualShapeProperties63.Append(nonVisualDrawingProperties77);
            nonVisualShapeProperties63.Append(nonVisualShapeDrawingProperties63);
            nonVisualShapeProperties63.Append(applicationNonVisualDrawingProperties77);

            ShapeProperties shapeProperties64 = new ShapeProperties();

            A.Transform2D transform2D28 = new A.Transform2D();
            A.Offset offset41 = new A.Offset(){ X = 3239493L, Y = 822856L };
            A.Extents extents41 = new A.Extents(){ Cx = 3857625L, Cy = 4061354L };

            transform2D28.Append(offset41);
            transform2D28.Append(extents41);

            shapeProperties64.Append(transform2D28);

            TextBody textBody63 = new TextBody();
            A.BodyProperties bodyProperties63 = new A.BodyProperties(){ Anchor = A.TextAnchoringTypeValues.Top };

            A.ListStyle listStyle63 = new A.ListStyle();

            A.Level1ParagraphProperties level1ParagraphProperties20 = new A.Level1ParagraphProperties(){ LeftMargin = 0, Indent = 0 };
            A.NoBullet noBullet61 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties97 = new A.DefaultRunProperties(){ FontSize = 2667 };

            level1ParagraphProperties20.Append(noBullet61);
            level1ParagraphProperties20.Append(defaultRunProperties97);

            A.Level2ParagraphProperties level2ParagraphProperties11 = new A.Level2ParagraphProperties(){ LeftMargin = 380985, Indent = 0 };
            A.NoBullet noBullet62 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties98 = new A.DefaultRunProperties(){ FontSize = 2333 };

            level2ParagraphProperties11.Append(noBullet62);
            level2ParagraphProperties11.Append(defaultRunProperties98);

            A.Level3ParagraphProperties level3ParagraphProperties11 = new A.Level3ParagraphProperties(){ LeftMargin = 761970, Indent = 0 };
            A.NoBullet noBullet63 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties99 = new A.DefaultRunProperties(){ FontSize = 2000 };

            level3ParagraphProperties11.Append(noBullet63);
            level3ParagraphProperties11.Append(defaultRunProperties99);

            A.Level4ParagraphProperties level4ParagraphProperties11 = new A.Level4ParagraphProperties(){ LeftMargin = 1142954, Indent = 0 };
            A.NoBullet noBullet64 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties100 = new A.DefaultRunProperties(){ FontSize = 1667 };

            level4ParagraphProperties11.Append(noBullet64);
            level4ParagraphProperties11.Append(defaultRunProperties100);

            A.Level5ParagraphProperties level5ParagraphProperties11 = new A.Level5ParagraphProperties(){ LeftMargin = 1523939, Indent = 0 };
            A.NoBullet noBullet65 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties101 = new A.DefaultRunProperties(){ FontSize = 1667 };

            level5ParagraphProperties11.Append(noBullet65);
            level5ParagraphProperties11.Append(defaultRunProperties101);

            A.Level6ParagraphProperties level6ParagraphProperties10 = new A.Level6ParagraphProperties(){ LeftMargin = 1904924, Indent = 0 };
            A.NoBullet noBullet66 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties102 = new A.DefaultRunProperties(){ FontSize = 1667 };

            level6ParagraphProperties10.Append(noBullet66);
            level6ParagraphProperties10.Append(defaultRunProperties102);

            A.Level7ParagraphProperties level7ParagraphProperties10 = new A.Level7ParagraphProperties(){ LeftMargin = 2285909, Indent = 0 };
            A.NoBullet noBullet67 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties103 = new A.DefaultRunProperties(){ FontSize = 1667 };

            level7ParagraphProperties10.Append(noBullet67);
            level7ParagraphProperties10.Append(defaultRunProperties103);

            A.Level8ParagraphProperties level8ParagraphProperties10 = new A.Level8ParagraphProperties(){ LeftMargin = 2666893, Indent = 0 };
            A.NoBullet noBullet68 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties104 = new A.DefaultRunProperties(){ FontSize = 1667 };

            level8ParagraphProperties10.Append(noBullet68);
            level8ParagraphProperties10.Append(defaultRunProperties104);

            A.Level9ParagraphProperties level9ParagraphProperties10 = new A.Level9ParagraphProperties(){ LeftMargin = 3047878, Indent = 0 };
            A.NoBullet noBullet69 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties105 = new A.DefaultRunProperties(){ FontSize = 1667 };

            level9ParagraphProperties10.Append(noBullet69);
            level9ParagraphProperties10.Append(defaultRunProperties105);

            listStyle63.Append(level1ParagraphProperties20);
            listStyle63.Append(level2ParagraphProperties11);
            listStyle63.Append(level3ParagraphProperties11);
            listStyle63.Append(level4ParagraphProperties11);
            listStyle63.Append(level5ParagraphProperties11);
            listStyle63.Append(level6ParagraphProperties10);
            listStyle63.Append(level7ParagraphProperties10);
            listStyle63.Append(level8ParagraphProperties10);
            listStyle63.Append(level9ParagraphProperties10);

            A.Paragraph paragraph99 = new A.Paragraph();

            A.Run run98 = new A.Run();
            A.RunProperties runProperties120 = new A.RunProperties(){ Language = "en-GB" };
            A.Text text120 = new A.Text();
            text120.Text = "Click icon to add picture";

            run98.Append(runProperties120);
            run98.Append(text120);
            A.EndParagraphRunProperties endParagraphRunProperties58 = new A.EndParagraphRunProperties(){ Language = "en-US", Dirty = false };

            paragraph99.Append(run98);
            paragraph99.Append(endParagraphRunProperties58);

            textBody63.Append(bodyProperties63);
            textBody63.Append(listStyle63);
            textBody63.Append(paragraph99);

            shape63.Append(nonVisualShapeProperties63);
            shape63.Append(shapeProperties64);
            shape63.Append(textBody63);

            Shape shape64 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties64 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties78 = new NonVisualDrawingProperties(){ Id = (UInt32Value)4U, Name = "Text Placeholder 3" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties64 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks63 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties64.Append(shapeLocks63);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties78 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape63 = new PlaceholderShape(){ Type = PlaceholderValues.Body, Size = PlaceholderSizeValues.Half, Index = (UInt32Value)2U };

            applicationNonVisualDrawingProperties78.Append(placeholderShape63);

            nonVisualShapeProperties64.Append(nonVisualDrawingProperties78);
            nonVisualShapeProperties64.Append(nonVisualShapeDrawingProperties64);
            nonVisualShapeProperties64.Append(applicationNonVisualDrawingProperties78);

            ShapeProperties shapeProperties65 = new ShapeProperties();

            A.Transform2D transform2D29 = new A.Transform2D();
            A.Offset offset42 = new A.Offset(){ X = 524868L, Y = 1714500L };
            A.Extents extents42 = new A.Extents(){ Cx = 2457648L, Cy = 3176323L };

            transform2D29.Append(offset42);
            transform2D29.Append(extents42);

            shapeProperties65.Append(transform2D29);

            TextBody textBody64 = new TextBody();
            A.BodyProperties bodyProperties64 = new A.BodyProperties();

            A.ListStyle listStyle64 = new A.ListStyle();

            A.Level1ParagraphProperties level1ParagraphProperties21 = new A.Level1ParagraphProperties(){ LeftMargin = 0, Indent = 0 };
            A.NoBullet noBullet70 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties106 = new A.DefaultRunProperties(){ FontSize = 1333 };

            level1ParagraphProperties21.Append(noBullet70);
            level1ParagraphProperties21.Append(defaultRunProperties106);

            A.Level2ParagraphProperties level2ParagraphProperties12 = new A.Level2ParagraphProperties(){ LeftMargin = 380985, Indent = 0 };
            A.NoBullet noBullet71 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties107 = new A.DefaultRunProperties(){ FontSize = 1167 };

            level2ParagraphProperties12.Append(noBullet71);
            level2ParagraphProperties12.Append(defaultRunProperties107);

            A.Level3ParagraphProperties level3ParagraphProperties12 = new A.Level3ParagraphProperties(){ LeftMargin = 761970, Indent = 0 };
            A.NoBullet noBullet72 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties108 = new A.DefaultRunProperties(){ FontSize = 1000 };

            level3ParagraphProperties12.Append(noBullet72);
            level3ParagraphProperties12.Append(defaultRunProperties108);

            A.Level4ParagraphProperties level4ParagraphProperties12 = new A.Level4ParagraphProperties(){ LeftMargin = 1142954, Indent = 0 };
            A.NoBullet noBullet73 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties109 = new A.DefaultRunProperties(){ FontSize = 833 };

            level4ParagraphProperties12.Append(noBullet73);
            level4ParagraphProperties12.Append(defaultRunProperties109);

            A.Level5ParagraphProperties level5ParagraphProperties12 = new A.Level5ParagraphProperties(){ LeftMargin = 1523939, Indent = 0 };
            A.NoBullet noBullet74 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties110 = new A.DefaultRunProperties(){ FontSize = 833 };

            level5ParagraphProperties12.Append(noBullet74);
            level5ParagraphProperties12.Append(defaultRunProperties110);

            A.Level6ParagraphProperties level6ParagraphProperties11 = new A.Level6ParagraphProperties(){ LeftMargin = 1904924, Indent = 0 };
            A.NoBullet noBullet75 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties111 = new A.DefaultRunProperties(){ FontSize = 833 };

            level6ParagraphProperties11.Append(noBullet75);
            level6ParagraphProperties11.Append(defaultRunProperties111);

            A.Level7ParagraphProperties level7ParagraphProperties11 = new A.Level7ParagraphProperties(){ LeftMargin = 2285909, Indent = 0 };
            A.NoBullet noBullet76 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties112 = new A.DefaultRunProperties(){ FontSize = 833 };

            level7ParagraphProperties11.Append(noBullet76);
            level7ParagraphProperties11.Append(defaultRunProperties112);

            A.Level8ParagraphProperties level8ParagraphProperties11 = new A.Level8ParagraphProperties(){ LeftMargin = 2666893, Indent = 0 };
            A.NoBullet noBullet77 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties113 = new A.DefaultRunProperties(){ FontSize = 833 };

            level8ParagraphProperties11.Append(noBullet77);
            level8ParagraphProperties11.Append(defaultRunProperties113);

            A.Level9ParagraphProperties level9ParagraphProperties11 = new A.Level9ParagraphProperties(){ LeftMargin = 3047878, Indent = 0 };
            A.NoBullet noBullet78 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties114 = new A.DefaultRunProperties(){ FontSize = 833 };

            level9ParagraphProperties11.Append(noBullet78);
            level9ParagraphProperties11.Append(defaultRunProperties114);

            listStyle64.Append(level1ParagraphProperties21);
            listStyle64.Append(level2ParagraphProperties12);
            listStyle64.Append(level3ParagraphProperties12);
            listStyle64.Append(level4ParagraphProperties12);
            listStyle64.Append(level5ParagraphProperties12);
            listStyle64.Append(level6ParagraphProperties11);
            listStyle64.Append(level7ParagraphProperties11);
            listStyle64.Append(level8ParagraphProperties11);
            listStyle64.Append(level9ParagraphProperties11);

            A.Paragraph paragraph100 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties63 = new A.ParagraphProperties(){ Level = 0 };

            A.Run run99 = new A.Run();
            A.RunProperties runProperties121 = new A.RunProperties(){ Language = "en-GB" };
            A.Text text121 = new A.Text();
            text121.Text = "Click to edit Master text styles";

            run99.Append(runProperties121);
            run99.Append(text121);

            paragraph100.Append(paragraphProperties63);
            paragraph100.Append(run99);

            textBody64.Append(bodyProperties64);
            textBody64.Append(listStyle64);
            textBody64.Append(paragraph100);

            shape64.Append(nonVisualShapeProperties64);
            shape64.Append(shapeProperties65);
            shape64.Append(textBody64);

            Shape shape65 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties65 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties79 = new NonVisualDrawingProperties(){ Id = (UInt32Value)5U, Name = "Date Placeholder 4" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties65 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks64 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties65.Append(shapeLocks64);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties79 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape64 = new PlaceholderShape(){ Type = PlaceholderValues.DateAndTime, Size = PlaceholderSizeValues.Half, Index = (UInt32Value)10U };

            applicationNonVisualDrawingProperties79.Append(placeholderShape64);

            nonVisualShapeProperties65.Append(nonVisualDrawingProperties79);
            nonVisualShapeProperties65.Append(nonVisualShapeDrawingProperties65);
            nonVisualShapeProperties65.Append(applicationNonVisualDrawingProperties79);
            ShapeProperties shapeProperties66 = new ShapeProperties();

            TextBody textBody65 = new TextBody();
            A.BodyProperties bodyProperties65 = new A.BodyProperties();
            A.ListStyle listStyle65 = new A.ListStyle();

            A.Paragraph paragraph101 = new A.Paragraph();

            A.Field field23 = new A.Field(){ Id = "{41037BA7-1B07-40F5-B3F2-74BAE740EAB6}", Type = "datetime1" };

            A.RunProperties runProperties122 = new A.RunProperties(){ Language = "en-US" };
            runProperties122.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            A.Text text122 = new A.Text();
            text122.Text = "1/17/2023";

            field23.Append(runProperties122);
            field23.Append(text122);
            A.EndParagraphRunProperties endParagraphRunProperties59 = new A.EndParagraphRunProperties(){ Language = "en-US", Dirty = false };

            paragraph101.Append(field23);
            paragraph101.Append(endParagraphRunProperties59);

            textBody65.Append(bodyProperties65);
            textBody65.Append(listStyle65);
            textBody65.Append(paragraph101);

            shape65.Append(nonVisualShapeProperties65);
            shape65.Append(shapeProperties66);
            shape65.Append(textBody65);

            Shape shape66 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties66 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties80 = new NonVisualDrawingProperties(){ Id = (UInt32Value)6U, Name = "Footer Placeholder 5" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties66 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks65 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties66.Append(shapeLocks65);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties80 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape65 = new PlaceholderShape(){ Type = PlaceholderValues.Footer, Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)11U };

            applicationNonVisualDrawingProperties80.Append(placeholderShape65);

            nonVisualShapeProperties66.Append(nonVisualDrawingProperties80);
            nonVisualShapeProperties66.Append(nonVisualShapeDrawingProperties66);
            nonVisualShapeProperties66.Append(applicationNonVisualDrawingProperties80);
            ShapeProperties shapeProperties67 = new ShapeProperties();

            TextBody textBody66 = new TextBody();
            A.BodyProperties bodyProperties66 = new A.BodyProperties();
            A.ListStyle listStyle66 = new A.ListStyle();

            A.Paragraph paragraph102 = new A.Paragraph();

            A.Run run100 = new A.Run();
            A.RunProperties runProperties123 = new A.RunProperties(){ Language = "en-US" };
            A.Text text123 = new A.Text();
            text123.Text = "Commercial & Workout Details";

            run100.Append(runProperties123);
            run100.Append(text123);
            A.EndParagraphRunProperties endParagraphRunProperties60 = new A.EndParagraphRunProperties(){ Language = "en-US", Dirty = false };

            paragraph102.Append(run100);
            paragraph102.Append(endParagraphRunProperties60);

            textBody66.Append(bodyProperties66);
            textBody66.Append(listStyle66);
            textBody66.Append(paragraph102);

            shape66.Append(nonVisualShapeProperties66);
            shape66.Append(shapeProperties67);
            shape66.Append(textBody66);

            Shape shape67 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties67 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties81 = new NonVisualDrawingProperties(){ Id = (UInt32Value)7U, Name = "Slide Number Placeholder 6" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties67 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks66 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties67.Append(shapeLocks66);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties81 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape66 = new PlaceholderShape(){ Type = PlaceholderValues.SlideNumber, Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)12U };

            applicationNonVisualDrawingProperties81.Append(placeholderShape66);

            nonVisualShapeProperties67.Append(nonVisualDrawingProperties81);
            nonVisualShapeProperties67.Append(nonVisualShapeDrawingProperties67);
            nonVisualShapeProperties67.Append(applicationNonVisualDrawingProperties81);
            ShapeProperties shapeProperties68 = new ShapeProperties();

            TextBody textBody67 = new TextBody();
            A.BodyProperties bodyProperties67 = new A.BodyProperties();
            A.ListStyle listStyle67 = new A.ListStyle();

            A.Paragraph paragraph103 = new A.Paragraph();

            A.Run run101 = new A.Run();

            A.RunProperties runProperties124 = new A.RunProperties(){ Language = "en-US" };

            A.SolidFill solidFill50 = new A.SolidFill();
            A.SchemeColor schemeColor66 = new A.SchemeColor(){ Val = A.SchemeColorValues.Accent2 };

            solidFill50.Append(schemeColor66);

            runProperties124.Append(solidFill50);
            A.Text text124 = new A.Text();
            text124.Text = "|";

            run101.Append(runProperties124);
            run101.Append(text124);

            A.Run run102 = new A.Run();
            A.RunProperties runProperties125 = new A.RunProperties(){ Language = "en-US" };
            A.Text text125 = new A.Text();
            text125.Text = "";

            run102.Append(runProperties125);
            run102.Append(text125);

            A.Field field24 = new A.Field(){ Id = "{E4F84C54-E2A4-46FF-B5B0-8F7A23C41D82}", Type = "slidenum" };

            A.RunProperties runProperties126 = new A.RunProperties(){ Language = "cs-CZ" };
            runProperties126.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            A.ParagraphProperties paragraphProperties64 = new A.ParagraphProperties();
            A.Text text126 = new A.Text();
            text126.Text = "‹#›";

            field24.Append(runProperties126);
            field24.Append(paragraphProperties64);
            field24.Append(text126);
            A.EndParagraphRunProperties endParagraphRunProperties61 = new A.EndParagraphRunProperties(){ Language = "cs-CZ", Dirty = false };

            paragraph103.Append(run101);
            paragraph103.Append(run102);
            paragraph103.Append(field24);
            paragraph103.Append(endParagraphRunProperties61);

            textBody67.Append(bodyProperties67);
            textBody67.Append(listStyle67);
            textBody67.Append(paragraph103);

            shape67.Append(nonVisualShapeProperties67);
            shape67.Append(shapeProperties68);
            shape67.Append(textBody67);

            shapeTree13.Append(nonVisualGroupShapeProperties13);
            shapeTree13.Append(groupShapeProperties13);
            shapeTree13.Append(shape62);
            shapeTree13.Append(shape63);
            shapeTree13.Append(shape64);
            shapeTree13.Append(shape65);
            shapeTree13.Append(shape66);
            shapeTree13.Append(shape67);

            CommonSlideDataExtensionList commonSlideDataExtensionList13 = new CommonSlideDataExtensionList();

            CommonSlideDataExtension commonSlideDataExtension13 = new CommonSlideDataExtension(){ Uri = "{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}" };

            P14.CreationId creationId13 = new P14.CreationId(){ Val = (UInt32Value)583440247U };
            creationId13.AddNamespaceDeclaration("p14", "http://schemas.microsoft.com/office/powerpoint/2010/main");

            commonSlideDataExtension13.Append(creationId13);

            commonSlideDataExtensionList13.Append(commonSlideDataExtension13);

            commonSlideData13.Append(shapeTree13);
            commonSlideData13.Append(commonSlideDataExtensionList13);

            ColorMapOverride colorMapOverride12 = new ColorMapOverride();
            A.MasterColorMapping masterColorMapping12 = new A.MasterColorMapping();

            colorMapOverride12.Append(masterColorMapping12);

            slideLayout12.Append(commonSlideData13);
            slideLayout12.Append(colorMapOverride12);

            slideLayoutPart12.SlideLayout = slideLayout12;
        }

        // Generates content of part.
        private void GeneratePartContent(SlideMasterPart part)
        {
            SlideMaster slideMaster2 = new SlideMaster();
            slideMaster2.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            slideMaster2.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            slideMaster2.AddNamespaceDeclaration("p", "http://schemas.openxmlformats.org/presentationml/2006/main");

            CommonSlideData commonSlideData14 = new CommonSlideData();

            Background background2 = new Background();

            BackgroundStyleReference backgroundStyleReference2 = new BackgroundStyleReference(){ Index = (UInt32Value)1001U };
            A.SchemeColor schemeColor67 = new A.SchemeColor(){ Val = A.SchemeColorValues.Background1 };

            backgroundStyleReference2.Append(schemeColor67);

            background2.Append(backgroundStyleReference2);

            ShapeTree shapeTree14 = new ShapeTree();

            NonVisualGroupShapeProperties nonVisualGroupShapeProperties14 = new NonVisualGroupShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties82 = new NonVisualDrawingProperties(){ Id = (UInt32Value)1U, Name = "" };
            NonVisualGroupShapeDrawingProperties nonVisualGroupShapeDrawingProperties14 = new NonVisualGroupShapeDrawingProperties();
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties82 = new ApplicationNonVisualDrawingProperties();

            nonVisualGroupShapeProperties14.Append(nonVisualDrawingProperties82);
            nonVisualGroupShapeProperties14.Append(nonVisualGroupShapeDrawingProperties14);
            nonVisualGroupShapeProperties14.Append(applicationNonVisualDrawingProperties82);

            GroupShapeProperties groupShapeProperties14 = new GroupShapeProperties();

            A.TransformGroup transformGroup14 = new A.TransformGroup();
            A.Offset offset43 = new A.Offset(){ X = 0L, Y = 0L };
            A.Extents extents43 = new A.Extents(){ Cx = 0L, Cy = 0L };
            A.ChildOffset childOffset14 = new A.ChildOffset(){ X = 0L, Y = 0L };
            A.ChildExtents childExtents14 = new A.ChildExtents(){ Cx = 0L, Cy = 0L };

            transformGroup14.Append(offset43);
            transformGroup14.Append(extents43);
            transformGroup14.Append(childOffset14);
            transformGroup14.Append(childExtents14);

            groupShapeProperties14.Append(transformGroup14);

            Shape shape68 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties68 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties83 = new NonVisualDrawingProperties(){ Id = (UInt32Value)2U, Name = "Title Placeholder 1" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties68 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks67 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties68.Append(shapeLocks67);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties83 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape67 = new PlaceholderShape(){ Type = PlaceholderValues.Title };

            applicationNonVisualDrawingProperties83.Append(placeholderShape67);

            nonVisualShapeProperties68.Append(nonVisualDrawingProperties83);
            nonVisualShapeProperties68.Append(nonVisualShapeDrawingProperties68);
            nonVisualShapeProperties68.Append(applicationNonVisualDrawingProperties83);

            ShapeProperties shapeProperties69 = new ShapeProperties();

            A.Transform2D transform2D30 = new A.Transform2D();
            A.Offset offset44 = new A.Offset(){ X = 523875L, Y = 304272L };
            A.Extents extents44 = new A.Extents(){ Cx = 6572250L, Cy = 1104636L };

            transform2D30.Append(offset44);
            transform2D30.Append(extents44);

            A.PresetGeometry presetGeometry8 = new A.PresetGeometry(){ Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList8 = new A.AdjustValueList();

            presetGeometry8.Append(adjustValueList8);

            shapeProperties69.Append(transform2D30);
            shapeProperties69.Append(presetGeometry8);

            TextBody textBody68 = new TextBody();

            A.BodyProperties bodyProperties68 = new A.BodyProperties(){ Vertical = A.TextVerticalValues.Horizontal, LeftInset = 91440, TopInset = 45720, RightInset = 91440, BottomInset = 45720, RightToLeftColumns = false, Anchor = A.TextAnchoringTypeValues.Center };
            A.NormalAutoFit normalAutoFit6 = new A.NormalAutoFit();

            bodyProperties68.Append(normalAutoFit6);
            A.ListStyle listStyle68 = new A.ListStyle();

            A.Paragraph paragraph104 = new A.Paragraph();

            A.Run run103 = new A.Run();
            A.RunProperties runProperties127 = new A.RunProperties(){ Language = "en-GB" };
            A.Text text127 = new A.Text();
            text127.Text = "Click to edit Master title style";

            run103.Append(runProperties127);
            run103.Append(text127);
            A.EndParagraphRunProperties endParagraphRunProperties62 = new A.EndParagraphRunProperties(){ Language = "en-US", Dirty = false };

            paragraph104.Append(run103);
            paragraph104.Append(endParagraphRunProperties62);

            textBody68.Append(bodyProperties68);
            textBody68.Append(listStyle68);
            textBody68.Append(paragraph104);

            shape68.Append(nonVisualShapeProperties68);
            shape68.Append(shapeProperties69);
            shape68.Append(textBody68);

            Shape shape69 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties69 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties84 = new NonVisualDrawingProperties(){ Id = (UInt32Value)3U, Name = "Text Placeholder 2" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties69 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks68 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties69.Append(shapeLocks68);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties84 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape68 = new PlaceholderShape(){ Type = PlaceholderValues.Body, Index = (UInt32Value)1U };

            applicationNonVisualDrawingProperties84.Append(placeholderShape68);

            nonVisualShapeProperties69.Append(nonVisualDrawingProperties84);
            nonVisualShapeProperties69.Append(nonVisualShapeDrawingProperties69);
            nonVisualShapeProperties69.Append(applicationNonVisualDrawingProperties84);

            ShapeProperties shapeProperties70 = new ShapeProperties();

            A.Transform2D transform2D31 = new A.Transform2D();
            A.Offset offset45 = new A.Offset(){ X = 523875L, Y = 1521354L };
            A.Extents extents45 = new A.Extents(){ Cx = 6572250L, Cy = 3626115L };

            transform2D31.Append(offset45);
            transform2D31.Append(extents45);

            A.PresetGeometry presetGeometry9 = new A.PresetGeometry(){ Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList9 = new A.AdjustValueList();

            presetGeometry9.Append(adjustValueList9);

            shapeProperties70.Append(transform2D31);
            shapeProperties70.Append(presetGeometry9);

            TextBody textBody69 = new TextBody();

            A.BodyProperties bodyProperties69 = new A.BodyProperties(){ Vertical = A.TextVerticalValues.Horizontal, LeftInset = 91440, TopInset = 45720, RightInset = 91440, BottomInset = 45720, RightToLeftColumns = false };
            A.NormalAutoFit normalAutoFit7 = new A.NormalAutoFit();

            bodyProperties69.Append(normalAutoFit7);
            A.ListStyle listStyle69 = new A.ListStyle();

            A.Paragraph paragraph105 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties65 = new A.ParagraphProperties(){ Level = 0 };

            A.Run run104 = new A.Run();
            A.RunProperties runProperties128 = new A.RunProperties(){ Language = "en-GB" };
            A.Text text128 = new A.Text();
            text128.Text = "Click to edit Master text styles";

            run104.Append(runProperties128);
            run104.Append(text128);

            paragraph105.Append(paragraphProperties65);
            paragraph105.Append(run104);

            A.Paragraph paragraph106 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties66 = new A.ParagraphProperties(){ Level = 1 };

            A.Run run105 = new A.Run();
            A.RunProperties runProperties129 = new A.RunProperties(){ Language = "en-GB" };
            A.Text text129 = new A.Text();
            text129.Text = "Second level";

            run105.Append(runProperties129);
            run105.Append(text129);

            paragraph106.Append(paragraphProperties66);
            paragraph106.Append(run105);

            A.Paragraph paragraph107 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties67 = new A.ParagraphProperties(){ Level = 2 };

            A.Run run106 = new A.Run();
            A.RunProperties runProperties130 = new A.RunProperties(){ Language = "en-GB" };
            A.Text text130 = new A.Text();
            text130.Text = "Third level";

            run106.Append(runProperties130);
            run106.Append(text130);

            paragraph107.Append(paragraphProperties67);
            paragraph107.Append(run106);

            A.Paragraph paragraph108 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties68 = new A.ParagraphProperties(){ Level = 3 };

            A.Run run107 = new A.Run();
            A.RunProperties runProperties131 = new A.RunProperties(){ Language = "en-GB" };
            A.Text text131 = new A.Text();
            text131.Text = "Fourth level";

            run107.Append(runProperties131);
            run107.Append(text131);

            paragraph108.Append(paragraphProperties68);
            paragraph108.Append(run107);

            A.Paragraph paragraph109 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties69 = new A.ParagraphProperties(){ Level = 4 };

            A.Run run108 = new A.Run();
            A.RunProperties runProperties132 = new A.RunProperties(){ Language = "en-GB" };
            A.Text text132 = new A.Text();
            text132.Text = "Fifth level";

            run108.Append(runProperties132);
            run108.Append(text132);
            A.EndParagraphRunProperties endParagraphRunProperties63 = new A.EndParagraphRunProperties(){ Language = "en-US", Dirty = false };

            paragraph109.Append(paragraphProperties69);
            paragraph109.Append(run108);
            paragraph109.Append(endParagraphRunProperties63);

            textBody69.Append(bodyProperties69);
            textBody69.Append(listStyle69);
            textBody69.Append(paragraph105);
            textBody69.Append(paragraph106);
            textBody69.Append(paragraph107);
            textBody69.Append(paragraph108);
            textBody69.Append(paragraph109);

            shape69.Append(nonVisualShapeProperties69);
            shape69.Append(shapeProperties70);
            shape69.Append(textBody69);

            Shape shape70 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties70 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties85 = new NonVisualDrawingProperties(){ Id = (UInt32Value)4U, Name = "Date Placeholder 3" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties70 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks69 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties70.Append(shapeLocks69);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties85 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape69 = new PlaceholderShape(){ Type = PlaceholderValues.DateAndTime, Size = PlaceholderSizeValues.Half, Index = (UInt32Value)2U };

            applicationNonVisualDrawingProperties85.Append(placeholderShape69);

            nonVisualShapeProperties70.Append(nonVisualDrawingProperties85);
            nonVisualShapeProperties70.Append(nonVisualShapeDrawingProperties70);
            nonVisualShapeProperties70.Append(applicationNonVisualDrawingProperties85);

            ShapeProperties shapeProperties71 = new ShapeProperties();

            A.Transform2D transform2D32 = new A.Transform2D();
            A.Offset offset46 = new A.Offset(){ X = 523875L, Y = 5296960L };
            A.Extents extents46 = new A.Extents(){ Cx = 1714500L, Cy = 304271L };

            transform2D32.Append(offset46);
            transform2D32.Append(extents46);

            A.PresetGeometry presetGeometry10 = new A.PresetGeometry(){ Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList10 = new A.AdjustValueList();

            presetGeometry10.Append(adjustValueList10);

            shapeProperties71.Append(transform2D32);
            shapeProperties71.Append(presetGeometry10);

            TextBody textBody70 = new TextBody();
            A.BodyProperties bodyProperties70 = new A.BodyProperties(){ Vertical = A.TextVerticalValues.Horizontal, LeftInset = 91440, TopInset = 45720, RightInset = 91440, BottomInset = 45720, RightToLeftColumns = false, Anchor = A.TextAnchoringTypeValues.Center };

            A.ListStyle listStyle70 = new A.ListStyle();

            A.Level1ParagraphProperties level1ParagraphProperties22 = new A.Level1ParagraphProperties(){ Alignment = A.TextAlignmentTypeValues.Left };

            A.DefaultRunProperties defaultRunProperties115 = new A.DefaultRunProperties(){ FontSize = 1000 };

            A.SolidFill solidFill51 = new A.SolidFill();

            A.SchemeColor schemeColor68 = new A.SchemeColor(){ Val = A.SchemeColorValues.Text1 };
            A.Tint tint19 = new A.Tint(){ Val = 75000 };

            schemeColor68.Append(tint19);

            solidFill51.Append(schemeColor68);

            defaultRunProperties115.Append(solidFill51);

            level1ParagraphProperties22.Append(defaultRunProperties115);

            listStyle70.Append(level1ParagraphProperties22);

            A.Paragraph paragraph110 = new A.Paragraph();

            A.Field field25 = new A.Field(){ Id = "{C5F68010-134A-445E-BD29-BEAAA1D9860C}", Type = "datetime1" };

            A.RunProperties runProperties133 = new A.RunProperties(){ Language = "en-US" };
            runProperties133.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            A.Text text133 = new A.Text();
            text133.Text = "1/17/2023";

            field25.Append(runProperties133);
            field25.Append(text133);
            A.EndParagraphRunProperties endParagraphRunProperties64 = new A.EndParagraphRunProperties(){ Language = "en-US", Dirty = false };

            paragraph110.Append(field25);
            paragraph110.Append(endParagraphRunProperties64);

            textBody70.Append(bodyProperties70);
            textBody70.Append(listStyle70);
            textBody70.Append(paragraph110);

            shape70.Append(nonVisualShapeProperties70);
            shape70.Append(shapeProperties71);
            shape70.Append(textBody70);

            Shape shape71 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties71 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties86 = new NonVisualDrawingProperties(){ Id = (UInt32Value)5U, Name = "Footer Placeholder 4" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties71 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks70 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties71.Append(shapeLocks70);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties86 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape70 = new PlaceholderShape(){ Type = PlaceholderValues.Footer, Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)3U };

            applicationNonVisualDrawingProperties86.Append(placeholderShape70);

            nonVisualShapeProperties71.Append(nonVisualDrawingProperties86);
            nonVisualShapeProperties71.Append(nonVisualShapeDrawingProperties71);
            nonVisualShapeProperties71.Append(applicationNonVisualDrawingProperties86);

            ShapeProperties shapeProperties72 = new ShapeProperties();

            A.Transform2D transform2D33 = new A.Transform2D();
            A.Offset offset47 = new A.Offset(){ X = 2524125L, Y = 5296960L };
            A.Extents extents47 = new A.Extents(){ Cx = 2571750L, Cy = 304271L };

            transform2D33.Append(offset47);
            transform2D33.Append(extents47);

            A.PresetGeometry presetGeometry11 = new A.PresetGeometry(){ Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList11 = new A.AdjustValueList();

            presetGeometry11.Append(adjustValueList11);

            shapeProperties72.Append(transform2D33);
            shapeProperties72.Append(presetGeometry11);

            TextBody textBody71 = new TextBody();
            A.BodyProperties bodyProperties71 = new A.BodyProperties(){ Vertical = A.TextVerticalValues.Horizontal, LeftInset = 91440, TopInset = 45720, RightInset = 91440, BottomInset = 45720, RightToLeftColumns = false, Anchor = A.TextAnchoringTypeValues.Center };

            A.ListStyle listStyle71 = new A.ListStyle();

            A.Level1ParagraphProperties level1ParagraphProperties23 = new A.Level1ParagraphProperties(){ Alignment = A.TextAlignmentTypeValues.Center };

            A.DefaultRunProperties defaultRunProperties116 = new A.DefaultRunProperties(){ FontSize = 1000 };

            A.SolidFill solidFill52 = new A.SolidFill();

            A.SchemeColor schemeColor69 = new A.SchemeColor(){ Val = A.SchemeColorValues.Text1 };
            A.Tint tint20 = new A.Tint(){ Val = 75000 };

            schemeColor69.Append(tint20);

            solidFill52.Append(schemeColor69);

            defaultRunProperties116.Append(solidFill52);

            level1ParagraphProperties23.Append(defaultRunProperties116);

            listStyle71.Append(level1ParagraphProperties23);

            A.Paragraph paragraph111 = new A.Paragraph();

            A.Run run109 = new A.Run();
            A.RunProperties runProperties134 = new A.RunProperties(){ Language = "en-US" };
            A.Text text134 = new A.Text();
            text134.Text = "Commercial & Workout Details";

            run109.Append(runProperties134);
            run109.Append(text134);
            A.EndParagraphRunProperties endParagraphRunProperties65 = new A.EndParagraphRunProperties(){ Language = "en-US", Dirty = false };

            paragraph111.Append(run109);
            paragraph111.Append(endParagraphRunProperties65);

            textBody71.Append(bodyProperties71);
            textBody71.Append(listStyle71);
            textBody71.Append(paragraph111);

            shape71.Append(nonVisualShapeProperties71);
            shape71.Append(shapeProperties72);
            shape71.Append(textBody71);

            Shape shape72 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties72 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties87 = new NonVisualDrawingProperties(){ Id = (UInt32Value)6U, Name = "Slide Number Placeholder 5" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties72 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks71 = new A.ShapeLocks(){ NoGrouping = true };

            nonVisualShapeDrawingProperties72.Append(shapeLocks71);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties87 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape71 = new PlaceholderShape(){ Type = PlaceholderValues.SlideNumber, Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)4U };

            applicationNonVisualDrawingProperties87.Append(placeholderShape71);

            nonVisualShapeProperties72.Append(nonVisualDrawingProperties87);
            nonVisualShapeProperties72.Append(nonVisualShapeDrawingProperties72);
            nonVisualShapeProperties72.Append(applicationNonVisualDrawingProperties87);

            ShapeProperties shapeProperties73 = new ShapeProperties();

            A.Transform2D transform2D34 = new A.Transform2D();
            A.Offset offset48 = new A.Offset(){ X = 5381625L, Y = 5296960L };
            A.Extents extents48 = new A.Extents(){ Cx = 1714500L, Cy = 304271L };

            transform2D34.Append(offset48);
            transform2D34.Append(extents48);

            A.PresetGeometry presetGeometry12 = new A.PresetGeometry(){ Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList12 = new A.AdjustValueList();

            presetGeometry12.Append(adjustValueList12);

            shapeProperties73.Append(transform2D34);
            shapeProperties73.Append(presetGeometry12);

            TextBody textBody72 = new TextBody();
            A.BodyProperties bodyProperties72 = new A.BodyProperties(){ Vertical = A.TextVerticalValues.Horizontal, LeftInset = 91440, TopInset = 45720, RightInset = 91440, BottomInset = 45720, RightToLeftColumns = false, Anchor = A.TextAnchoringTypeValues.Center };

            A.ListStyle listStyle72 = new A.ListStyle();

            A.Level1ParagraphProperties level1ParagraphProperties24 = new A.Level1ParagraphProperties(){ Alignment = A.TextAlignmentTypeValues.Right };

            A.DefaultRunProperties defaultRunProperties117 = new A.DefaultRunProperties(){ FontSize = 1000 };

            A.SolidFill solidFill53 = new A.SolidFill();

            A.SchemeColor schemeColor70 = new A.SchemeColor(){ Val = A.SchemeColorValues.Text1 };
            A.Tint tint21 = new A.Tint(){ Val = 75000 };

            schemeColor70.Append(tint21);

            solidFill53.Append(schemeColor70);

            defaultRunProperties117.Append(solidFill53);

            level1ParagraphProperties24.Append(defaultRunProperties117);

            listStyle72.Append(level1ParagraphProperties24);

            A.Paragraph paragraph112 = new A.Paragraph();

            A.Run run110 = new A.Run();

            A.RunProperties runProperties135 = new A.RunProperties(){ Language = "en-US" };

            A.SolidFill solidFill54 = new A.SolidFill();
            A.SchemeColor schemeColor71 = new A.SchemeColor(){ Val = A.SchemeColorValues.Accent2 };

            solidFill54.Append(schemeColor71);

            runProperties135.Append(solidFill54);
            A.Text text135 = new A.Text();
            text135.Text = "|";

            run110.Append(runProperties135);
            run110.Append(text135);

            A.Run run111 = new A.Run();
            A.RunProperties runProperties136 = new A.RunProperties(){ Language = "en-US" };
            A.Text text136 = new A.Text();
            text136.Text = "";

            run111.Append(runProperties136);
            run111.Append(text136);

            A.Field field26 = new A.Field(){ Id = "{E4F84C54-E2A4-46FF-B5B0-8F7A23C41D82}", Type = "slidenum" };

            A.RunProperties runProperties137 = new A.RunProperties(){ Language = "cs-CZ" };
            runProperties137.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            A.ParagraphProperties paragraphProperties70 = new A.ParagraphProperties();
            A.Text text137 = new A.Text();
            text137.Text = "‹#›";

            field26.Append(runProperties137);
            field26.Append(paragraphProperties70);
            field26.Append(text137);
            A.EndParagraphRunProperties endParagraphRunProperties66 = new A.EndParagraphRunProperties(){ Language = "cs-CZ", Dirty = false };

            paragraph112.Append(run110);
            paragraph112.Append(run111);
            paragraph112.Append(field26);
            paragraph112.Append(endParagraphRunProperties66);

            textBody72.Append(bodyProperties72);
            textBody72.Append(listStyle72);
            textBody72.Append(paragraph112);

            shape72.Append(nonVisualShapeProperties72);
            shape72.Append(shapeProperties73);
            shape72.Append(textBody72);

            Shape shape73 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties73 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties88 = new NonVisualDrawingProperties(){ Id = (UInt32Value)7U, Name = "Rectangle 6" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList2 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension2 = new A.NonVisualDrawingPropertiesExtension(){ Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            OpenXmlUnknownElement openXmlUnknownElement2 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{DABD3684-197E-4B2B-99DA-0FFE27DA5C10}\" />");

            nonVisualDrawingPropertiesExtension2.Append(openXmlUnknownElement2);

            nonVisualDrawingPropertiesExtensionList2.Append(nonVisualDrawingPropertiesExtension2);

            nonVisualDrawingProperties88.Append(nonVisualDrawingPropertiesExtensionList2);
            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties73 = new NonVisualShapeDrawingProperties();
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties88 = new ApplicationNonVisualDrawingProperties(){ UserDrawn = true };

            nonVisualShapeProperties73.Append(nonVisualDrawingProperties88);
            nonVisualShapeProperties73.Append(nonVisualShapeDrawingProperties73);
            nonVisualShapeProperties73.Append(applicationNonVisualDrawingProperties88);

            ShapeProperties shapeProperties74 = new ShapeProperties();

            A.Transform2D transform2D35 = new A.Transform2D();
            A.Offset offset49 = new A.Offset(){ X = 269874L, Y = 1075788L };
            A.Extents extents49 = new A.Extents(){ Cx = 7110678L, Cy = 36000L };

            transform2D35.Append(offset49);
            transform2D35.Append(extents49);

            A.PresetGeometry presetGeometry13 = new A.PresetGeometry(){ Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList13 = new A.AdjustValueList();

            presetGeometry13.Append(adjustValueList13);

            A.GradientFill gradientFill5 = new A.GradientFill(){ Flip = A.TileFlipValues.None, RotateWithShape = true };

            A.GradientStopList gradientStopList5 = new A.GradientStopList();

            A.GradientStop gradientStop12 = new A.GradientStop(){ Position = 0 };
            A.SchemeColor schemeColor72 = new A.SchemeColor(){ Val = A.SchemeColorValues.Accent2 };

            gradientStop12.Append(schemeColor72);

            A.GradientStop gradientStop13 = new A.GradientStop(){ Position = 100000 };
            A.SchemeColor schemeColor73 = new A.SchemeColor(){ Val = A.SchemeColorValues.Accent1 };

            gradientStop13.Append(schemeColor73);

            gradientStopList5.Append(gradientStop12);
            gradientStopList5.Append(gradientStop13);
            A.LinearGradientFill linearGradientFill5 = new A.LinearGradientFill(){ Angle = 0, Scaled = true };
            A.TileRectangle tileRectangle2 = new A.TileRectangle();

            gradientFill5.Append(gradientStopList5);
            gradientFill5.Append(linearGradientFill5);
            gradientFill5.Append(tileRectangle2);

            A.Outline outline5 = new A.Outline();
            A.NoFill noFill2 = new A.NoFill();

            outline5.Append(noFill2);

            shapeProperties74.Append(transform2D35);
            shapeProperties74.Append(presetGeometry13);
            shapeProperties74.Append(gradientFill5);
            shapeProperties74.Append(outline5);

            ShapeStyle shapeStyle2 = new ShapeStyle();

            A.LineReference lineReference2 = new A.LineReference(){ Index = (UInt32Value)2U };

            A.SchemeColor schemeColor74 = new A.SchemeColor(){ Val = A.SchemeColorValues.Accent1 };
            A.Shade shade7 = new A.Shade(){ Val = 50000 };

            schemeColor74.Append(shade7);

            lineReference2.Append(schemeColor74);

            A.FillReference fillReference2 = new A.FillReference(){ Index = (UInt32Value)1U };
            A.SchemeColor schemeColor75 = new A.SchemeColor(){ Val = A.SchemeColorValues.Accent1 };

            fillReference2.Append(schemeColor75);

            A.EffectReference effectReference2 = new A.EffectReference(){ Index = (UInt32Value)0U };
            A.SchemeColor schemeColor76 = new A.SchemeColor(){ Val = A.SchemeColorValues.Accent1 };

            effectReference2.Append(schemeColor76);

            A.FontReference fontReference2 = new A.FontReference(){ Index = A.FontCollectionIndexValues.Minor };
            A.SchemeColor schemeColor77 = new A.SchemeColor(){ Val = A.SchemeColorValues.Light1 };

            fontReference2.Append(schemeColor77);

            shapeStyle2.Append(lineReference2);
            shapeStyle2.Append(fillReference2);
            shapeStyle2.Append(effectReference2);
            shapeStyle2.Append(fontReference2);

            TextBody textBody73 = new TextBody();
            A.BodyProperties bodyProperties73 = new A.BodyProperties(){ RightToLeftColumns = false, Anchor = A.TextAnchoringTypeValues.Center };
            A.ListStyle listStyle73 = new A.ListStyle();

            A.Paragraph paragraph113 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties71 = new A.ParagraphProperties(){ Alignment = A.TextAlignmentTypeValues.Center };
            A.EndParagraphRunProperties endParagraphRunProperties67 = new A.EndParagraphRunProperties(){ Language = "cs-CZ", FontSize = 1404 };

            paragraph113.Append(paragraphProperties71);
            paragraph113.Append(endParagraphRunProperties67);

            textBody73.Append(bodyProperties73);
            textBody73.Append(listStyle73);
            textBody73.Append(paragraph113);

            shape73.Append(nonVisualShapeProperties73);
            shape73.Append(shapeProperties74);
            shape73.Append(shapeStyle2);
            shape73.Append(textBody73);

            shapeTree14.Append(nonVisualGroupShapeProperties14);
            shapeTree14.Append(groupShapeProperties14);
            shapeTree14.Append(shape68);
            shapeTree14.Append(shape69);
            shapeTree14.Append(shape70);
            shapeTree14.Append(shape71);
            shapeTree14.Append(shape72);
            shapeTree14.Append(shape73);

            CommonSlideDataExtensionList commonSlideDataExtensionList14 = new CommonSlideDataExtensionList();

            CommonSlideDataExtension commonSlideDataExtension14 = new CommonSlideDataExtension(){ Uri = "{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}" };

            P14.CreationId creationId14 = new P14.CreationId(){ Val = (UInt32Value)1735501897U };
            creationId14.AddNamespaceDeclaration("p14", "http://schemas.microsoft.com/office/powerpoint/2010/main");

            commonSlideDataExtension14.Append(creationId14);

            commonSlideDataExtensionList14.Append(commonSlideDataExtension14);

            commonSlideData14.Append(background2);
            commonSlideData14.Append(shapeTree14);
            commonSlideData14.Append(commonSlideDataExtensionList14);
            ColorMap colorMap2 = new ColorMap(){ Background1 = A.ColorSchemeIndexValues.Light1, Text1 = A.ColorSchemeIndexValues.Dark1, Background2 = A.ColorSchemeIndexValues.Light2, Text2 = A.ColorSchemeIndexValues.Dark2, Accent1 = A.ColorSchemeIndexValues.Accent1, Accent2 = A.ColorSchemeIndexValues.Accent2, Accent3 = A.ColorSchemeIndexValues.Accent3, Accent4 = A.ColorSchemeIndexValues.Accent4, Accent5 = A.ColorSchemeIndexValues.Accent5, Accent6 = A.ColorSchemeIndexValues.Accent6, Hyperlink = A.ColorSchemeIndexValues.Hyperlink, FollowedHyperlink = A.ColorSchemeIndexValues.FollowedHyperlink };

            SlideLayoutIdList slideLayoutIdList2 = new SlideLayoutIdList();
            SlideLayoutId slideLayoutId13 = new SlideLayoutId(){ Id = (UInt32Value)2147483663U, RelationshipId = "rId1" };
            SlideLayoutId slideLayoutId14 = new SlideLayoutId(){ Id = (UInt32Value)2147483664U, RelationshipId = "rId2" };
            SlideLayoutId slideLayoutId15 = new SlideLayoutId(){ Id = (UInt32Value)2147483665U, RelationshipId = "rId3" };
            SlideLayoutId slideLayoutId16 = new SlideLayoutId(){ Id = (UInt32Value)2147483666U, RelationshipId = "rId4" };
            SlideLayoutId slideLayoutId17 = new SlideLayoutId(){ Id = (UInt32Value)2147483667U, RelationshipId = "rId5" };
            SlideLayoutId slideLayoutId18 = new SlideLayoutId(){ Id = (UInt32Value)2147483668U, RelationshipId = "rId6" };
            SlideLayoutId slideLayoutId19 = new SlideLayoutId(){ Id = (UInt32Value)2147483669U, RelationshipId = "rId7" };
            SlideLayoutId slideLayoutId20 = new SlideLayoutId(){ Id = (UInt32Value)2147483670U, RelationshipId = "rId8" };
            SlideLayoutId slideLayoutId21 = new SlideLayoutId(){ Id = (UInt32Value)2147483671U, RelationshipId = "rId9" };
            SlideLayoutId slideLayoutId22 = new SlideLayoutId(){ Id = (UInt32Value)2147483672U, RelationshipId = "rId10" };
            SlideLayoutId slideLayoutId23 = new SlideLayoutId(){ Id = (UInt32Value)2147483673U, RelationshipId = "rId11" };
            SlideLayoutId slideLayoutId24 = new SlideLayoutId(){ Id = (UInt32Value)2147483676U, RelationshipId = "rId12" };

            slideLayoutIdList2.Append(slideLayoutId13);
            slideLayoutIdList2.Append(slideLayoutId14);
            slideLayoutIdList2.Append(slideLayoutId15);
            slideLayoutIdList2.Append(slideLayoutId16);
            slideLayoutIdList2.Append(slideLayoutId17);
            slideLayoutIdList2.Append(slideLayoutId18);
            slideLayoutIdList2.Append(slideLayoutId19);
            slideLayoutIdList2.Append(slideLayoutId20);
            slideLayoutIdList2.Append(slideLayoutId21);
            slideLayoutIdList2.Append(slideLayoutId22);
            slideLayoutIdList2.Append(slideLayoutId23);
            slideLayoutIdList2.Append(slideLayoutId24);
            HeaderFooter headerFooter2 = new HeaderFooter(){ Header = false, DateTime = false };

            TextStyles textStyles2 = new TextStyles();

            TitleStyle titleStyle2 = new TitleStyle();

            A.Level1ParagraphProperties level1ParagraphProperties25 = new A.Level1ParagraphProperties(){ Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 761970, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.LineSpacing lineSpacing11 = new A.LineSpacing();
            A.SpacingPercent spacingPercent12 = new A.SpacingPercent(){ Val = 90000 };

            lineSpacing11.Append(spacingPercent12);

            A.SpaceBefore spaceBefore11 = new A.SpaceBefore();
            A.SpacingPercent spacingPercent13 = new A.SpacingPercent(){ Val = 0 };

            spaceBefore11.Append(spacingPercent13);
            A.NoBullet noBullet79 = new A.NoBullet();

            A.DefaultRunProperties defaultRunProperties118 = new A.DefaultRunProperties(){ FontSize = 3667, Kerning = 1200 };

            A.SolidFill solidFill55 = new A.SolidFill();
            A.SchemeColor schemeColor78 = new A.SchemeColor(){ Val = A.SchemeColorValues.Text1 };

            solidFill55.Append(schemeColor78);
            A.LatinFont latinFont22 = new A.LatinFont(){ Typeface = "+mj-lt" };
            A.EastAsianFont eastAsianFont22 = new A.EastAsianFont(){ Typeface = "+mj-ea" };
            A.ComplexScriptFont complexScriptFont22 = new A.ComplexScriptFont(){ Typeface = "+mj-cs" };

            defaultRunProperties118.Append(solidFill55);
            defaultRunProperties118.Append(latinFont22);
            defaultRunProperties118.Append(eastAsianFont22);
            defaultRunProperties118.Append(complexScriptFont22);

            level1ParagraphProperties25.Append(lineSpacing11);
            level1ParagraphProperties25.Append(spaceBefore11);
            level1ParagraphProperties25.Append(noBullet79);
            level1ParagraphProperties25.Append(defaultRunProperties118);

            titleStyle2.Append(level1ParagraphProperties25);

            BodyStyle bodyStyle2 = new BodyStyle();

            A.Level1ParagraphProperties level1ParagraphProperties26 = new A.Level1ParagraphProperties(){ LeftMargin = 190492, Indent = -190492, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 761970, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.LineSpacing lineSpacing12 = new A.LineSpacing();
            A.SpacingPercent spacingPercent14 = new A.SpacingPercent(){ Val = 90000 };

            lineSpacing12.Append(spacingPercent14);

            A.SpaceBefore spaceBefore12 = new A.SpaceBefore();
            A.SpacingPoints spacingPoints10 = new A.SpacingPoints(){ Val = 833 };

            spaceBefore12.Append(spacingPoints10);
            A.BulletFont bulletFont12 = new A.BulletFont(){ Typeface = "Arial", Panose = "020B0604020202020204", PitchFamily = 34, CharacterSet = 0 };
            A.CharacterBullet characterBullet10 = new A.CharacterBullet(){ Char = "•" };

            A.DefaultRunProperties defaultRunProperties119 = new A.DefaultRunProperties(){ FontSize = 2333, Kerning = 1200 };

            A.SolidFill solidFill56 = new A.SolidFill();
            A.SchemeColor schemeColor79 = new A.SchemeColor(){ Val = A.SchemeColorValues.Text1 };

            solidFill56.Append(schemeColor79);
            A.LatinFont latinFont23 = new A.LatinFont(){ Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont23 = new A.EastAsianFont(){ Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont23 = new A.ComplexScriptFont(){ Typeface = "+mn-cs" };

            defaultRunProperties119.Append(solidFill56);
            defaultRunProperties119.Append(latinFont23);
            defaultRunProperties119.Append(eastAsianFont23);
            defaultRunProperties119.Append(complexScriptFont23);

            level1ParagraphProperties26.Append(lineSpacing12);
            level1ParagraphProperties26.Append(spaceBefore12);
            level1ParagraphProperties26.Append(bulletFont12);
            level1ParagraphProperties26.Append(characterBullet10);
            level1ParagraphProperties26.Append(defaultRunProperties119);

            A.Level2ParagraphProperties level2ParagraphProperties13 = new A.Level2ParagraphProperties(){ LeftMargin = 571477, Indent = -190492, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 761970, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.LineSpacing lineSpacing13 = new A.LineSpacing();
            A.SpacingPercent spacingPercent15 = new A.SpacingPercent(){ Val = 90000 };

            lineSpacing13.Append(spacingPercent15);

            A.SpaceBefore spaceBefore13 = new A.SpaceBefore();
            A.SpacingPoints spacingPoints11 = new A.SpacingPoints(){ Val = 417 };

            spaceBefore13.Append(spacingPoints11);
            A.BulletFont bulletFont13 = new A.BulletFont(){ Typeface = "Arial", Panose = "020B0604020202020204", PitchFamily = 34, CharacterSet = 0 };
            A.CharacterBullet characterBullet11 = new A.CharacterBullet(){ Char = "•" };

            A.DefaultRunProperties defaultRunProperties120 = new A.DefaultRunProperties(){ FontSize = 2000, Kerning = 1200 };

            A.SolidFill solidFill57 = new A.SolidFill();
            A.SchemeColor schemeColor80 = new A.SchemeColor(){ Val = A.SchemeColorValues.Text1 };

            solidFill57.Append(schemeColor80);
            A.LatinFont latinFont24 = new A.LatinFont(){ Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont24 = new A.EastAsianFont(){ Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont24 = new A.ComplexScriptFont(){ Typeface = "+mn-cs" };

            defaultRunProperties120.Append(solidFill57);
            defaultRunProperties120.Append(latinFont24);
            defaultRunProperties120.Append(eastAsianFont24);
            defaultRunProperties120.Append(complexScriptFont24);

            level2ParagraphProperties13.Append(lineSpacing13);
            level2ParagraphProperties13.Append(spaceBefore13);
            level2ParagraphProperties13.Append(bulletFont13);
            level2ParagraphProperties13.Append(characterBullet11);
            level2ParagraphProperties13.Append(defaultRunProperties120);

            A.Level3ParagraphProperties level3ParagraphProperties13 = new A.Level3ParagraphProperties(){ LeftMargin = 952462, Indent = -190492, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 761970, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.LineSpacing lineSpacing14 = new A.LineSpacing();
            A.SpacingPercent spacingPercent16 = new A.SpacingPercent(){ Val = 90000 };

            lineSpacing14.Append(spacingPercent16);

            A.SpaceBefore spaceBefore14 = new A.SpaceBefore();
            A.SpacingPoints spacingPoints12 = new A.SpacingPoints(){ Val = 417 };

            spaceBefore14.Append(spacingPoints12);
            A.BulletFont bulletFont14 = new A.BulletFont(){ Typeface = "Arial", Panose = "020B0604020202020204", PitchFamily = 34, CharacterSet = 0 };
            A.CharacterBullet characterBullet12 = new A.CharacterBullet(){ Char = "•" };

            A.DefaultRunProperties defaultRunProperties121 = new A.DefaultRunProperties(){ FontSize = 1667, Kerning = 1200 };

            A.SolidFill solidFill58 = new A.SolidFill();
            A.SchemeColor schemeColor81 = new A.SchemeColor(){ Val = A.SchemeColorValues.Text1 };

            solidFill58.Append(schemeColor81);
            A.LatinFont latinFont25 = new A.LatinFont(){ Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont25 = new A.EastAsianFont(){ Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont25 = new A.ComplexScriptFont(){ Typeface = "+mn-cs" };

            defaultRunProperties121.Append(solidFill58);
            defaultRunProperties121.Append(latinFont25);
            defaultRunProperties121.Append(eastAsianFont25);
            defaultRunProperties121.Append(complexScriptFont25);

            level3ParagraphProperties13.Append(lineSpacing14);
            level3ParagraphProperties13.Append(spaceBefore14);
            level3ParagraphProperties13.Append(bulletFont14);
            level3ParagraphProperties13.Append(characterBullet12);
            level3ParagraphProperties13.Append(defaultRunProperties121);

            A.Level4ParagraphProperties level4ParagraphProperties13 = new A.Level4ParagraphProperties(){ LeftMargin = 1333447, Indent = -190492, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 761970, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.LineSpacing lineSpacing15 = new A.LineSpacing();
            A.SpacingPercent spacingPercent17 = new A.SpacingPercent(){ Val = 90000 };

            lineSpacing15.Append(spacingPercent17);

            A.SpaceBefore spaceBefore15 = new A.SpaceBefore();
            A.SpacingPoints spacingPoints13 = new A.SpacingPoints(){ Val = 417 };

            spaceBefore15.Append(spacingPoints13);
            A.BulletFont bulletFont15 = new A.BulletFont(){ Typeface = "Arial", Panose = "020B0604020202020204", PitchFamily = 34, CharacterSet = 0 };
            A.CharacterBullet characterBullet13 = new A.CharacterBullet(){ Char = "•" };

            A.DefaultRunProperties defaultRunProperties122 = new A.DefaultRunProperties(){ FontSize = 1500, Kerning = 1200 };

            A.SolidFill solidFill59 = new A.SolidFill();
            A.SchemeColor schemeColor82 = new A.SchemeColor(){ Val = A.SchemeColorValues.Text1 };

            solidFill59.Append(schemeColor82);
            A.LatinFont latinFont26 = new A.LatinFont(){ Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont26 = new A.EastAsianFont(){ Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont26 = new A.ComplexScriptFont(){ Typeface = "+mn-cs" };

            defaultRunProperties122.Append(solidFill59);
            defaultRunProperties122.Append(latinFont26);
            defaultRunProperties122.Append(eastAsianFont26);
            defaultRunProperties122.Append(complexScriptFont26);

            level4ParagraphProperties13.Append(lineSpacing15);
            level4ParagraphProperties13.Append(spaceBefore15);
            level4ParagraphProperties13.Append(bulletFont15);
            level4ParagraphProperties13.Append(characterBullet13);
            level4ParagraphProperties13.Append(defaultRunProperties122);

            A.Level5ParagraphProperties level5ParagraphProperties13 = new A.Level5ParagraphProperties(){ LeftMargin = 1714431, Indent = -190492, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 761970, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.LineSpacing lineSpacing16 = new A.LineSpacing();
            A.SpacingPercent spacingPercent18 = new A.SpacingPercent(){ Val = 90000 };

            lineSpacing16.Append(spacingPercent18);

            A.SpaceBefore spaceBefore16 = new A.SpaceBefore();
            A.SpacingPoints spacingPoints14 = new A.SpacingPoints(){ Val = 417 };

            spaceBefore16.Append(spacingPoints14);
            A.BulletFont bulletFont16 = new A.BulletFont(){ Typeface = "Arial", Panose = "020B0604020202020204", PitchFamily = 34, CharacterSet = 0 };
            A.CharacterBullet characterBullet14 = new A.CharacterBullet(){ Char = "•" };

            A.DefaultRunProperties defaultRunProperties123 = new A.DefaultRunProperties(){ FontSize = 1500, Kerning = 1200 };

            A.SolidFill solidFill60 = new A.SolidFill();
            A.SchemeColor schemeColor83 = new A.SchemeColor(){ Val = A.SchemeColorValues.Text1 };

            solidFill60.Append(schemeColor83);
            A.LatinFont latinFont27 = new A.LatinFont(){ Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont27 = new A.EastAsianFont(){ Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont27 = new A.ComplexScriptFont(){ Typeface = "+mn-cs" };

            defaultRunProperties123.Append(solidFill60);
            defaultRunProperties123.Append(latinFont27);
            defaultRunProperties123.Append(eastAsianFont27);
            defaultRunProperties123.Append(complexScriptFont27);

            level5ParagraphProperties13.Append(lineSpacing16);
            level5ParagraphProperties13.Append(spaceBefore16);
            level5ParagraphProperties13.Append(bulletFont16);
            level5ParagraphProperties13.Append(characterBullet14);
            level5ParagraphProperties13.Append(defaultRunProperties123);

            A.Level6ParagraphProperties level6ParagraphProperties12 = new A.Level6ParagraphProperties(){ LeftMargin = 2095416, Indent = -190492, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 761970, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.LineSpacing lineSpacing17 = new A.LineSpacing();
            A.SpacingPercent spacingPercent19 = new A.SpacingPercent(){ Val = 90000 };

            lineSpacing17.Append(spacingPercent19);

            A.SpaceBefore spaceBefore17 = new A.SpaceBefore();
            A.SpacingPoints spacingPoints15 = new A.SpacingPoints(){ Val = 417 };

            spaceBefore17.Append(spacingPoints15);
            A.BulletFont bulletFont17 = new A.BulletFont(){ Typeface = "Arial", Panose = "020B0604020202020204", PitchFamily = 34, CharacterSet = 0 };
            A.CharacterBullet characterBullet15 = new A.CharacterBullet(){ Char = "•" };

            A.DefaultRunProperties defaultRunProperties124 = new A.DefaultRunProperties(){ FontSize = 1500, Kerning = 1200 };

            A.SolidFill solidFill61 = new A.SolidFill();
            A.SchemeColor schemeColor84 = new A.SchemeColor(){ Val = A.SchemeColorValues.Text1 };

            solidFill61.Append(schemeColor84);
            A.LatinFont latinFont28 = new A.LatinFont(){ Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont28 = new A.EastAsianFont(){ Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont28 = new A.ComplexScriptFont(){ Typeface = "+mn-cs" };

            defaultRunProperties124.Append(solidFill61);
            defaultRunProperties124.Append(latinFont28);
            defaultRunProperties124.Append(eastAsianFont28);
            defaultRunProperties124.Append(complexScriptFont28);

            level6ParagraphProperties12.Append(lineSpacing17);
            level6ParagraphProperties12.Append(spaceBefore17);
            level6ParagraphProperties12.Append(bulletFont17);
            level6ParagraphProperties12.Append(characterBullet15);
            level6ParagraphProperties12.Append(defaultRunProperties124);

            A.Level7ParagraphProperties level7ParagraphProperties12 = new A.Level7ParagraphProperties(){ LeftMargin = 2476401, Indent = -190492, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 761970, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.LineSpacing lineSpacing18 = new A.LineSpacing();
            A.SpacingPercent spacingPercent20 = new A.SpacingPercent(){ Val = 90000 };

            lineSpacing18.Append(spacingPercent20);

            A.SpaceBefore spaceBefore18 = new A.SpaceBefore();
            A.SpacingPoints spacingPoints16 = new A.SpacingPoints(){ Val = 417 };

            spaceBefore18.Append(spacingPoints16);
            A.BulletFont bulletFont18 = new A.BulletFont(){ Typeface = "Arial", Panose = "020B0604020202020204", PitchFamily = 34, CharacterSet = 0 };
            A.CharacterBullet characterBullet16 = new A.CharacterBullet(){ Char = "•" };

            A.DefaultRunProperties defaultRunProperties125 = new A.DefaultRunProperties(){ FontSize = 1500, Kerning = 1200 };

            A.SolidFill solidFill62 = new A.SolidFill();
            A.SchemeColor schemeColor85 = new A.SchemeColor(){ Val = A.SchemeColorValues.Text1 };

            solidFill62.Append(schemeColor85);
            A.LatinFont latinFont29 = new A.LatinFont(){ Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont29 = new A.EastAsianFont(){ Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont29 = new A.ComplexScriptFont(){ Typeface = "+mn-cs" };

            defaultRunProperties125.Append(solidFill62);
            defaultRunProperties125.Append(latinFont29);
            defaultRunProperties125.Append(eastAsianFont29);
            defaultRunProperties125.Append(complexScriptFont29);

            level7ParagraphProperties12.Append(lineSpacing18);
            level7ParagraphProperties12.Append(spaceBefore18);
            level7ParagraphProperties12.Append(bulletFont18);
            level7ParagraphProperties12.Append(characterBullet16);
            level7ParagraphProperties12.Append(defaultRunProperties125);

            A.Level8ParagraphProperties level8ParagraphProperties12 = new A.Level8ParagraphProperties(){ LeftMargin = 2857386, Indent = -190492, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 761970, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.LineSpacing lineSpacing19 = new A.LineSpacing();
            A.SpacingPercent spacingPercent21 = new A.SpacingPercent(){ Val = 90000 };

            lineSpacing19.Append(spacingPercent21);

            A.SpaceBefore spaceBefore19 = new A.SpaceBefore();
            A.SpacingPoints spacingPoints17 = new A.SpacingPoints(){ Val = 417 };

            spaceBefore19.Append(spacingPoints17);
            A.BulletFont bulletFont19 = new A.BulletFont(){ Typeface = "Arial", Panose = "020B0604020202020204", PitchFamily = 34, CharacterSet = 0 };
            A.CharacterBullet characterBullet17 = new A.CharacterBullet(){ Char = "•" };

            A.DefaultRunProperties defaultRunProperties126 = new A.DefaultRunProperties(){ FontSize = 1500, Kerning = 1200 };

            A.SolidFill solidFill63 = new A.SolidFill();
            A.SchemeColor schemeColor86 = new A.SchemeColor(){ Val = A.SchemeColorValues.Text1 };

            solidFill63.Append(schemeColor86);
            A.LatinFont latinFont30 = new A.LatinFont(){ Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont30 = new A.EastAsianFont(){ Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont30 = new A.ComplexScriptFont(){ Typeface = "+mn-cs" };

            defaultRunProperties126.Append(solidFill63);
            defaultRunProperties126.Append(latinFont30);
            defaultRunProperties126.Append(eastAsianFont30);
            defaultRunProperties126.Append(complexScriptFont30);

            level8ParagraphProperties12.Append(lineSpacing19);
            level8ParagraphProperties12.Append(spaceBefore19);
            level8ParagraphProperties12.Append(bulletFont19);
            level8ParagraphProperties12.Append(characterBullet17);
            level8ParagraphProperties12.Append(defaultRunProperties126);

            A.Level9ParagraphProperties level9ParagraphProperties12 = new A.Level9ParagraphProperties(){ LeftMargin = 3238370, Indent = -190492, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 761970, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.LineSpacing lineSpacing20 = new A.LineSpacing();
            A.SpacingPercent spacingPercent22 = new A.SpacingPercent(){ Val = 90000 };

            lineSpacing20.Append(spacingPercent22);

            A.SpaceBefore spaceBefore20 = new A.SpaceBefore();
            A.SpacingPoints spacingPoints18 = new A.SpacingPoints(){ Val = 417 };

            spaceBefore20.Append(spacingPoints18);
            A.BulletFont bulletFont20 = new A.BulletFont(){ Typeface = "Arial", Panose = "020B0604020202020204", PitchFamily = 34, CharacterSet = 0 };
            A.CharacterBullet characterBullet18 = new A.CharacterBullet(){ Char = "•" };

            A.DefaultRunProperties defaultRunProperties127 = new A.DefaultRunProperties(){ FontSize = 1500, Kerning = 1200 };

            A.SolidFill solidFill64 = new A.SolidFill();
            A.SchemeColor schemeColor87 = new A.SchemeColor(){ Val = A.SchemeColorValues.Text1 };

            solidFill64.Append(schemeColor87);
            A.LatinFont latinFont31 = new A.LatinFont(){ Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont31 = new A.EastAsianFont(){ Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont31 = new A.ComplexScriptFont(){ Typeface = "+mn-cs" };

            defaultRunProperties127.Append(solidFill64);
            defaultRunProperties127.Append(latinFont31);
            defaultRunProperties127.Append(eastAsianFont31);
            defaultRunProperties127.Append(complexScriptFont31);

            level9ParagraphProperties12.Append(lineSpacing20);
            level9ParagraphProperties12.Append(spaceBefore20);
            level9ParagraphProperties12.Append(bulletFont20);
            level9ParagraphProperties12.Append(characterBullet18);
            level9ParagraphProperties12.Append(defaultRunProperties127);

            bodyStyle2.Append(level1ParagraphProperties26);
            bodyStyle2.Append(level2ParagraphProperties13);
            bodyStyle2.Append(level3ParagraphProperties13);
            bodyStyle2.Append(level4ParagraphProperties13);
            bodyStyle2.Append(level5ParagraphProperties13);
            bodyStyle2.Append(level6ParagraphProperties12);
            bodyStyle2.Append(level7ParagraphProperties12);
            bodyStyle2.Append(level8ParagraphProperties12);
            bodyStyle2.Append(level9ParagraphProperties12);

            OtherStyle otherStyle2 = new OtherStyle();

            A.DefaultParagraphProperties defaultParagraphProperties2 = new A.DefaultParagraphProperties();
            A.DefaultRunProperties defaultRunProperties128 = new A.DefaultRunProperties(){ Language = "en-US" };

            defaultParagraphProperties2.Append(defaultRunProperties128);

            A.Level1ParagraphProperties level1ParagraphProperties27 = new A.Level1ParagraphProperties(){ LeftMargin = 0, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 761970, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.DefaultRunProperties defaultRunProperties129 = new A.DefaultRunProperties(){ FontSize = 1500, Kerning = 1200 };

            A.SolidFill solidFill65 = new A.SolidFill();
            A.SchemeColor schemeColor88 = new A.SchemeColor(){ Val = A.SchemeColorValues.Text1 };

            solidFill65.Append(schemeColor88);
            A.LatinFont latinFont32 = new A.LatinFont(){ Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont32 = new A.EastAsianFont(){ Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont32 = new A.ComplexScriptFont(){ Typeface = "+mn-cs" };

            defaultRunProperties129.Append(solidFill65);
            defaultRunProperties129.Append(latinFont32);
            defaultRunProperties129.Append(eastAsianFont32);
            defaultRunProperties129.Append(complexScriptFont32);

            level1ParagraphProperties27.Append(defaultRunProperties129);

            A.Level2ParagraphProperties level2ParagraphProperties14 = new A.Level2ParagraphProperties(){ LeftMargin = 380985, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 761970, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.DefaultRunProperties defaultRunProperties130 = new A.DefaultRunProperties(){ FontSize = 1500, Kerning = 1200 };

            A.SolidFill solidFill66 = new A.SolidFill();
            A.SchemeColor schemeColor89 = new A.SchemeColor(){ Val = A.SchemeColorValues.Text1 };

            solidFill66.Append(schemeColor89);
            A.LatinFont latinFont33 = new A.LatinFont(){ Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont33 = new A.EastAsianFont(){ Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont33 = new A.ComplexScriptFont(){ Typeface = "+mn-cs" };

            defaultRunProperties130.Append(solidFill66);
            defaultRunProperties130.Append(latinFont33);
            defaultRunProperties130.Append(eastAsianFont33);
            defaultRunProperties130.Append(complexScriptFont33);

            level2ParagraphProperties14.Append(defaultRunProperties130);

            A.Level3ParagraphProperties level3ParagraphProperties14 = new A.Level3ParagraphProperties(){ LeftMargin = 761970, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 761970, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.DefaultRunProperties defaultRunProperties131 = new A.DefaultRunProperties(){ FontSize = 1500, Kerning = 1200 };

            A.SolidFill solidFill67 = new A.SolidFill();
            A.SchemeColor schemeColor90 = new A.SchemeColor(){ Val = A.SchemeColorValues.Text1 };

            solidFill67.Append(schemeColor90);
            A.LatinFont latinFont34 = new A.LatinFont(){ Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont34 = new A.EastAsianFont(){ Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont34 = new A.ComplexScriptFont(){ Typeface = "+mn-cs" };

            defaultRunProperties131.Append(solidFill67);
            defaultRunProperties131.Append(latinFont34);
            defaultRunProperties131.Append(eastAsianFont34);
            defaultRunProperties131.Append(complexScriptFont34);

            level3ParagraphProperties14.Append(defaultRunProperties131);

            A.Level4ParagraphProperties level4ParagraphProperties14 = new A.Level4ParagraphProperties(){ LeftMargin = 1142954, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 761970, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.DefaultRunProperties defaultRunProperties132 = new A.DefaultRunProperties(){ FontSize = 1500, Kerning = 1200 };

            A.SolidFill solidFill68 = new A.SolidFill();
            A.SchemeColor schemeColor91 = new A.SchemeColor(){ Val = A.SchemeColorValues.Text1 };

            solidFill68.Append(schemeColor91);
            A.LatinFont latinFont35 = new A.LatinFont(){ Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont35 = new A.EastAsianFont(){ Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont35 = new A.ComplexScriptFont(){ Typeface = "+mn-cs" };

            defaultRunProperties132.Append(solidFill68);
            defaultRunProperties132.Append(latinFont35);
            defaultRunProperties132.Append(eastAsianFont35);
            defaultRunProperties132.Append(complexScriptFont35);

            level4ParagraphProperties14.Append(defaultRunProperties132);

            A.Level5ParagraphProperties level5ParagraphProperties14 = new A.Level5ParagraphProperties(){ LeftMargin = 1523939, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 761970, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.DefaultRunProperties defaultRunProperties133 = new A.DefaultRunProperties(){ FontSize = 1500, Kerning = 1200 };

            A.SolidFill solidFill69 = new A.SolidFill();
            A.SchemeColor schemeColor92 = new A.SchemeColor(){ Val = A.SchemeColorValues.Text1 };

            solidFill69.Append(schemeColor92);
            A.LatinFont latinFont36 = new A.LatinFont(){ Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont36 = new A.EastAsianFont(){ Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont36 = new A.ComplexScriptFont(){ Typeface = "+mn-cs" };

            defaultRunProperties133.Append(solidFill69);
            defaultRunProperties133.Append(latinFont36);
            defaultRunProperties133.Append(eastAsianFont36);
            defaultRunProperties133.Append(complexScriptFont36);

            level5ParagraphProperties14.Append(defaultRunProperties133);

            A.Level6ParagraphProperties level6ParagraphProperties13 = new A.Level6ParagraphProperties(){ LeftMargin = 1904924, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 761970, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.DefaultRunProperties defaultRunProperties134 = new A.DefaultRunProperties(){ FontSize = 1500, Kerning = 1200 };

            A.SolidFill solidFill70 = new A.SolidFill();
            A.SchemeColor schemeColor93 = new A.SchemeColor(){ Val = A.SchemeColorValues.Text1 };

            solidFill70.Append(schemeColor93);
            A.LatinFont latinFont37 = new A.LatinFont(){ Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont37 = new A.EastAsianFont(){ Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont37 = new A.ComplexScriptFont(){ Typeface = "+mn-cs" };

            defaultRunProperties134.Append(solidFill70);
            defaultRunProperties134.Append(latinFont37);
            defaultRunProperties134.Append(eastAsianFont37);
            defaultRunProperties134.Append(complexScriptFont37);

            level6ParagraphProperties13.Append(defaultRunProperties134);

            A.Level7ParagraphProperties level7ParagraphProperties13 = new A.Level7ParagraphProperties(){ LeftMargin = 2285909, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 761970, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.DefaultRunProperties defaultRunProperties135 = new A.DefaultRunProperties(){ FontSize = 1500, Kerning = 1200 };

            A.SolidFill solidFill71 = new A.SolidFill();
            A.SchemeColor schemeColor94 = new A.SchemeColor(){ Val = A.SchemeColorValues.Text1 };

            solidFill71.Append(schemeColor94);
            A.LatinFont latinFont38 = new A.LatinFont(){ Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont38 = new A.EastAsianFont(){ Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont38 = new A.ComplexScriptFont(){ Typeface = "+mn-cs" };

            defaultRunProperties135.Append(solidFill71);
            defaultRunProperties135.Append(latinFont38);
            defaultRunProperties135.Append(eastAsianFont38);
            defaultRunProperties135.Append(complexScriptFont38);

            level7ParagraphProperties13.Append(defaultRunProperties135);

            A.Level8ParagraphProperties level8ParagraphProperties13 = new A.Level8ParagraphProperties(){ LeftMargin = 2666893, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 761970, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.DefaultRunProperties defaultRunProperties136 = new A.DefaultRunProperties(){ FontSize = 1500, Kerning = 1200 };

            A.SolidFill solidFill72 = new A.SolidFill();
            A.SchemeColor schemeColor95 = new A.SchemeColor(){ Val = A.SchemeColorValues.Text1 };

            solidFill72.Append(schemeColor95);
            A.LatinFont latinFont39 = new A.LatinFont(){ Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont39 = new A.EastAsianFont(){ Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont39 = new A.ComplexScriptFont(){ Typeface = "+mn-cs" };

            defaultRunProperties136.Append(solidFill72);
            defaultRunProperties136.Append(latinFont39);
            defaultRunProperties136.Append(eastAsianFont39);
            defaultRunProperties136.Append(complexScriptFont39);

            level8ParagraphProperties13.Append(defaultRunProperties136);

            A.Level9ParagraphProperties level9ParagraphProperties13 = new A.Level9ParagraphProperties(){ LeftMargin = 3047878, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 761970, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.DefaultRunProperties defaultRunProperties137 = new A.DefaultRunProperties(){ FontSize = 1500, Kerning = 1200 };

            A.SolidFill solidFill73 = new A.SolidFill();
            A.SchemeColor schemeColor96 = new A.SchemeColor(){ Val = A.SchemeColorValues.Text1 };

            solidFill73.Append(schemeColor96);
            A.LatinFont latinFont40 = new A.LatinFont(){ Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont40 = new A.EastAsianFont(){ Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont40 = new A.ComplexScriptFont(){ Typeface = "+mn-cs" };

            defaultRunProperties137.Append(solidFill73);
            defaultRunProperties137.Append(latinFont40);
            defaultRunProperties137.Append(eastAsianFont40);
            defaultRunProperties137.Append(complexScriptFont40);

            level9ParagraphProperties13.Append(defaultRunProperties137);

            otherStyle2.Append(defaultParagraphProperties2);
            otherStyle2.Append(level1ParagraphProperties27);
            otherStyle2.Append(level2ParagraphProperties14);
            otherStyle2.Append(level3ParagraphProperties14);
            otherStyle2.Append(level4ParagraphProperties14);
            otherStyle2.Append(level5ParagraphProperties14);
            otherStyle2.Append(level6ParagraphProperties13);
            otherStyle2.Append(level7ParagraphProperties13);
            otherStyle2.Append(level8ParagraphProperties13);
            otherStyle2.Append(level9ParagraphProperties13);

            textStyles2.Append(titleStyle2);
            textStyles2.Append(bodyStyle2);
            textStyles2.Append(otherStyle2);

            SlideMasterExtensionList slideMasterExtensionList2 = new SlideMasterExtensionList();

            SlideMasterExtension slideMasterExtension2 = new SlideMasterExtension(){ Uri = "{27BBF7A9-308A-43DC-89C8-2F10F3537804}" };

            P15.SlideGuideList slideGuideList2 = new P15.SlideGuideList();
            slideGuideList2.AddNamespaceDeclaration("p15", "http://schemas.microsoft.com/office/powerpoint/2012/main");

            P15.ExtendedGuide extendedGuide9 = new P15.ExtendedGuide(){ Id = (UInt32Value)1U, Orientation = DirectionValues.Horizontal, Position = 99 };

            P15.ColorType colorType9 = new P15.ColorType();
            A.RgbColorModelHex rgbColorModelHex20 = new A.RgbColorModelHex(){ Val = "F26B43" };

            colorType9.Append(rgbColorModelHex20);

            extendedGuide9.Append(colorType9);

            P15.ExtendedGuide extendedGuide10 = new P15.ExtendedGuide(){ Id = (UInt32Value)2U, Position = 170 };

            P15.ColorType colorType10 = new P15.ColorType();
            A.RgbColorModelHex rgbColorModelHex21 = new A.RgbColorModelHex(){ Val = "F26B43" };

            colorType10.Append(rgbColorModelHex21);

            extendedGuide10.Append(colorType10);

            P15.ExtendedGuide extendedGuide11 = new P15.ExtendedGuide(){ Id = (UInt32Value)3U, Position = 4649 };

            P15.ColorType colorType11 = new P15.ColorType();
            A.RgbColorModelHex rgbColorModelHex22 = new A.RgbColorModelHex(){ Val = "F26B43" };

            colorType11.Append(rgbColorModelHex22);

            extendedGuide11.Append(colorType11);

            P15.ExtendedGuide extendedGuide12 = new P15.ExtendedGuide(){ Id = (UInt32Value)4U, Orientation = DirectionValues.Horizontal, Position = 689 };

            P15.ColorType colorType12 = new P15.ColorType();
            A.RgbColorModelHex rgbColorModelHex23 = new A.RgbColorModelHex(){ Val = "F26B43" };

            colorType12.Append(rgbColorModelHex23);

            extendedGuide12.Append(colorType12);

            P15.ExtendedGuide extendedGuide13 = new P15.ExtendedGuide(){ Id = (UInt32Value)5U, Orientation = DirectionValues.Horizontal, Position = 839 };

            P15.ColorType colorType13 = new P15.ColorType();
            A.RgbColorModelHex rgbColorModelHex24 = new A.RgbColorModelHex(){ Val = "F26B43" };

            colorType13.Append(rgbColorModelHex24);

            extendedGuide13.Append(colorType13);

            P15.ExtendedGuide extendedGuide14 = new P15.ExtendedGuide(){ Id = (UInt32Value)6U, Orientation = DirectionValues.Horizontal, Position = 3161 };

            P15.ColorType colorType14 = new P15.ColorType();
            A.RgbColorModelHex rgbColorModelHex25 = new A.RgbColorModelHex(){ Val = "F26B43" };

            colorType14.Append(rgbColorModelHex25);

            extendedGuide14.Append(colorType14);

            P15.ExtendedGuide extendedGuide15 = new P15.ExtendedGuide(){ Id = (UInt32Value)7U, Orientation = DirectionValues.Horizontal, Position = 3342 };

            P15.ColorType colorType15 = new P15.ColorType();
            A.RgbColorModelHex rgbColorModelHex26 = new A.RgbColorModelHex(){ Val = "F26B43" };

            colorType15.Append(rgbColorModelHex26);

            extendedGuide15.Append(colorType15);

            P15.ExtendedGuide extendedGuide16 = new P15.ExtendedGuide(){ Id = (UInt32Value)8U, Orientation = DirectionValues.Horizontal, Position = 3546 };

            P15.ColorType colorType16 = new P15.ColorType();
            A.RgbColorModelHex rgbColorModelHex27 = new A.RgbColorModelHex(){ Val = "F26B43" };

            colorType16.Append(rgbColorModelHex27);

            extendedGuide16.Append(colorType16);

            slideGuideList2.Append(extendedGuide9);
            slideGuideList2.Append(extendedGuide10);
            slideGuideList2.Append(extendedGuide11);
            slideGuideList2.Append(extendedGuide12);
            slideGuideList2.Append(extendedGuide13);
            slideGuideList2.Append(extendedGuide14);
            slideGuideList2.Append(extendedGuide15);
            slideGuideList2.Append(extendedGuide16);

            slideMasterExtension2.Append(slideGuideList2);

            slideMasterExtensionList2.Append(slideMasterExtension2);

            slideMaster2.Append(commonSlideData14);
            slideMaster2.Append(colorMap2);
            slideMaster2.Append(slideLayoutIdList2);
            slideMaster2.Append(headerFooter2);
            slideMaster2.Append(textStyles2);
            slideMaster2.Append(slideMasterExtensionList2);

            part.SlideMaster = slideMaster2;
        }

        #region Binary Data
        private string imagePart1Data = "/9j/4AAQSkZJRgABAQEA3ADcAAD/2wBDAAIBAQEBAQIBAQECAgICAgQDAgICAgUEBAMEBgUGBgYFBgYGBwkIBgcJBwYGCAsICQoKCgoKBggLDAsKDAkKCgr/2wBDAQICAgICAgUDAwUKBwYHCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgr/wAARCAEQAlEDASIAAhEBAxEB/8QAHwAAAQUBAQEBAQEAAAAAAAAAAAECAwQFBgcICQoL/8QAtRAAAgEDAwIEAwUFBAQAAAF9AQIDAAQRBRIhMUEGE1FhByJxFDKBkaEII0KxwRVS0fAkM2JyggkKFhcYGRolJicoKSo0NTY3ODk6Q0RFRkdISUpTVFVWV1hZWmNkZWZnaGlqc3R1dnd4eXqDhIWGh4iJipKTlJWWl5iZmqKjpKWmp6ipqrKztLW2t7i5usLDxMXGx8jJytLT1NXW19jZ2uHi4+Tl5ufo6erx8vP09fb3+Pn6/8QAHwEAAwEBAQEBAQEBAQAAAAAAAAECAwQFBgcICQoL/8QAtREAAgECBAQDBAcFBAQAAQJ3AAECAxEEBSExBhJBUQdhcRMiMoEIFEKRobHBCSMzUvAVYnLRChYkNOEl8RcYGRomJygpKjU2Nzg5OkNERUZHSElKU1RVVldYWVpjZGVmZ2hpanN0dXZ3eHl6goOEhYaHiImKkpOUlZaXmJmaoqOkpaanqKmqsrO0tba3uLm6wsPExcbHyMnK0tPU1dbX2Nna4uPk5ebn6Onq8vP09fb3+Pn6/9oADAMBAAIRAxEAPwD9/KKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooozQAUUUUAFFFFABRRRQAUUUUAFFFFABRRnHWjNABRRRQAUUUUAFFGaKACiiigAoozRnPSgAooooAKKKKACiiigAooozQAUUUUAFFFGaACijOOtFABR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";

        private System.IO.Stream GetBinaryDataStream(string base64String)
        {
            return new System.IO.MemoryStream(System.Convert.FromBase64String(base64String));
        }

        #endregion

    }
}

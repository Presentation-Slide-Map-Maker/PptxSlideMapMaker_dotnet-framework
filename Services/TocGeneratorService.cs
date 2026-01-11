using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using A = DocumentFormat.OpenXml.Drawing;
using Application = Microsoft.Office.Interop.PowerPoint.Application;
using Presentation = Microsoft.Office.Interop.PowerPoint.Presentation;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;
using Slide = Microsoft.Office.Interop.PowerPoint.Slide;

namespace TocBuilder_dotnet_framework.Services
{
    internal class TocLinkInfo
    {
        public string ShapeGuid { get; set; }
        public int SlideNumber { get; set; }
    }
    public class TocGeneratorService : IDisposable
    {
        private bool _disposed;
        private readonly List<TocLinkInfo> _pendingLinks = new List<TocLinkInfo>();

        public string CreateTableOfContents(
            string inputPath,
            List<Models.SlideItem> slides,
            int columns,
            int margin,
            int backgroundSlideIndex)
        {
            Application pptApp = null;
            Presentation pres = null;

            string outputPath = GetOutputPath(inputPath);

            // 1. Копируем исходный файл
            File.Copy(inputPath, outputPath, true);

            // 2. Временная папка для миниатюр
            string tempDir = Path.Combine(
                Path.GetTempPath(),
                "toc_thumbs_" + Guid.NewGuid().ToString("N"));

            Directory.CreateDirectory(tempDir);

            try
            {
                // 3. Interpop для экспорта слайдов в PNG
                pptApp = new Application { DisplayAlerts = PpAlertLevel.ppAlertsNone };
                pres = pptApp.Presentations.Open(
                    outputPath,
                    MsoTriState.msoFalse,
                    MsoTriState.msoFalse,
                    MsoTriState.msoFalse);

                foreach (var slide in slides)
                {
                    pres.Slides[slide.Number].Export(
                        Path.Combine(tempDir, $"slide_{slide.Number}.png"),
                        "PNG",
                        1600,
                        900);
                }

                pres.Close();
                Marshal.ReleaseComObject(pres);
                pres = null;

                pptApp.Quit();
                Marshal.ReleaseComObject(pptApp);
                pptApp = null;

                GC.Collect();
                GC.WaitForPendingFinalizers();

                // 4. OpenXml для создания содержания
                BuildTocSlideOpenXml(
                    outputPath,
                    slides,
                    columns,
                    margin,
                    tempDir,
                    backgroundSlideIndex);

                return outputPath;
            }
            finally
            {
                TryDeleteFolder(tempDir);
            }
        }

        private void BuildTocSlideOpenXml(
            string pptxPath,
            List<Models.SlideItem> slides,
            int columns,
            int margin,
            string tempDir,
            int backgroundSlideIndex
        )
        {
            using (var doc = PresentationDocument.Open(pptxPath, true))
            {
                // Указание авторства
                var props = doc.PackageProperties;
                props.Creator = "TOC Builder";
                props.LastModifiedBy = "TOC Builder";


                var presPart = doc.PresentationPart;
                var slideIds = presPart.Presentation.SlideIdList.Elements<SlideId>().ToList();

                // layout
                if (backgroundSlideIndex < 0 || backgroundSlideIndex >= slideIds.Count)
                    throw new ArgumentOutOfRangeException(nameof(backgroundSlideIndex));

                var backgroundSlideId = slideIds[backgroundSlideIndex];
                var backgroundSlidePart =
                    (SlidePart)presPart.GetPartById(backgroundSlideId.RelationshipId);
                var layoutPart = backgroundSlidePart.SlideLayoutPart;


                // TOC
                var tocPart = presPart.AddNewPart<SlidePart>();
                tocPart.Slide = new DocumentFormat.OpenXml.Presentation.Slide(
                    new CommonSlideData(new ShapeTree()),
                    new ColorMapOverride(new A.MasterColorMapping()));

                var notesPart = tocPart.AddNewPart<NotesSlidePart>(); // Упоминание в заметках
                notesPart.NotesSlide = new NotesSlide(
                    new CommonSlideData(
                        new ShapeTree(
                            new NonVisualGroupShapeProperties(
                                new NonVisualDrawingProperties { Id = 1, Name = "" },
                                new NonVisualGroupShapeDrawingProperties(),
                                new ApplicationNonVisualDrawingProperties()
                            ),
                            new GroupShapeProperties(new A.TransformGroup()),

                            new DocumentFormat.OpenXml.Presentation.Shape(
                                new NonVisualShapeProperties(
                                    new NonVisualDrawingProperties
                                    {
                                        Id = 2,
                                        Name = "Notes Placeholder"
                                    },
                                    new NonVisualShapeDrawingProperties(
                                        new A.ShapeLocks { NoGrouping = true }
                                    ),
                                    new ApplicationNonVisualDrawingProperties(
                                        new PlaceholderShape
                                        {
                                            Type = PlaceholderValues.Body
                                        }
                                    )
                                ),
                                new ShapeProperties(),
                                new TextBody(
                                    new A.BodyProperties(),
                                    new A.ListStyle(),
                                    new A.Paragraph(
                                        new A.Run(
                                            new A.Text("Made with TOC Builder")
                                        )
                                    )
                                )
                            )
                        )
                    )
                );
                notesPart.AddPart(presPart.NotesMasterPart); // привязка к NotesMaster


                tocPart.AddPart(layoutPart);
                tocPart.Slide.Save();

                presPart.Presentation.SlideIdList.Append(
                    new SlideId
                    {
                        Id = slideIds.Max(s => s.Id.Value) + 1,
                        RelationshipId = presPart.GetIdOfPart(tocPart)
                    });

                var shapeTree = tocPart.Slide.CommonSlideData.ShapeTree;

                float slideWidthPx = presPart.Presentation.SlideSize.Cx / 9525f;
                float slideHeightPx = presPart.Presentation.SlideSize.Cy / 9525f;

                var layout = LayoutCalculatorService.CalculateOptimalLayout(
                    slides.Count,
                    margin,
                    columns,
                    slideWidthPx,
                    slideHeightPx
                );


                for (int i = 0; i < slides.Count; i++)
                {
                    int row = i / layout.Columns;
                    int col = i % layout.Columns;

                    long x = (long)((margin + col * (layout.ThumbWidth + margin)) * 9525);
                    long y = (long)((LayoutConstants.TitleHeight +
                                     row * layout.RowHeight) * 9525);

                    long thumbW = (long)(layout.ThumbWidth * 9525);
                    long thumbH = (long)(layout.ThumbHeight * 9525);
                    long captionH = (long)(LayoutConstants.CaptionHeight * 9525);


                    var targetSlidePart =
                        (SlidePart)presPart.GetPartById(slideIds[slides[i].Number - 1].RelationshipId);

                    tocPart.AddPart(targetSlidePart);
                    string relId = tocPart.GetIdOfPart(targetSlidePart);

                    var imgPart = tocPart.AddImagePart(ImagePartType.Png);
                    using (var fs = File.OpenRead(Path.Combine(tempDir, $"slide_{slides[i].Number}.png")))
                        imgPart.FeedData(fs);

                    // Рамка
                    shapeTree.Append(
                        new DocumentFormat.OpenXml.Presentation.Shape(
                            new NonVisualShapeProperties(
                                new NonVisualDrawingProperties
                                {
                                    Id = (uint)(3000 + i),
                                    Name = $"Frame {slides[i].Number}"
                                },
                                new NonVisualShapeDrawingProperties(),
                                new ApplicationNonVisualDrawingProperties()
                            ),
                            new ShapeProperties(
                                new A.Transform2D(
                                    new A.Offset { X = x, Y = y },
                                    new A.Extents { Cx = thumbW, Cy = thumbH }),
                                new A.PresetGeometry(new A.AdjustValueList())
                                { Preset = A.ShapeTypeValues.Rectangle },
                                new A.Outline(
                                    new A.SolidFill(
                                        new A.RgbColorModelHex { Val = "808080" }
                                    )
                                //,new A.Width { Val = 12700 } // ~1.5pt
                                ),
                                new A.NoFill()
                            )
                        )
                    );

                    // Миниатюра
                    shapeTree.Append(
                        new Picture(
                            new NonVisualPictureProperties(
                                new NonVisualDrawingProperties
                                {
                                    Id = (uint)(2000 + i),
                                    Name = $"Slide {slides[i].Number}",
                                    HyperlinkOnClick = new A.HyperlinkOnClick
                                    {
                                        Id = relId,
                                        Action = "ppaction://hlinksldjump"
                                    }
                                },
                                new NonVisualPictureDrawingProperties(),
                                new ApplicationNonVisualDrawingProperties()
                            ),
                            new BlipFill(
                                new A.Blip { Embed = tocPart.GetIdOfPart(imgPart) },
                                new A.Stretch(new A.FillRectangle())
                            ),
                            new ShapeProperties(
                                new A.Transform2D(
                                    new A.Offset { X = x, Y = y },
                                    new A.Extents { Cx = thumbW, Cy = thumbH }),
                                new A.PresetGeometry(new A.AdjustValueList())
                                { Preset = A.ShapeTypeValues.Rectangle })
                        ));


                    // Подпись
                    long captionHeight = 300000;
                    long captionY = y + thumbH;

                    shapeTree.Append(
                        new DocumentFormat.OpenXml.Presentation.Shape(
                            new NonVisualShapeProperties(
                                new NonVisualDrawingProperties
                                {
                                    Id = (uint)(4000 + i),
                                    Name = $"Caption {slides[i].Number}"
                                },
                                new NonVisualShapeDrawingProperties(),
                                new ApplicationNonVisualDrawingProperties()
                            ),
                            new ShapeProperties(
                                new A.Transform2D(
                                    new A.Offset { X = x, Y = captionY },
                                    new A.Extents { Cx = thumbW, Cy = captionHeight }),
                                new A.PresetGeometry(new A.AdjustValueList())
                                { Preset = A.ShapeTypeValues.Rectangle }
                            ),
                            new TextBody(
                                new A.BodyProperties(),
                                new A.ListStyle(),
                                new A.Paragraph(
                                    new A.Run(
                                        new A.RunProperties(
                                            new A.HyperlinkOnClick
                                            {
                                                Id = relId,
                                                Action = "ppaction://hlinksldjump"
                                            })
                                        {
                                            Language = "ru-RU"
                                        },
                                        new A.Text($"Слайд {slides[i].Number}")
                                    )
                                )
                            )
                        )
                    );


                }

                // Создание обратной навигации после добавления TOC-слайда в SlideIdList
                string tocRelId = presPart.GetIdOfPart(tocPart);
                int totalSlides = presPart.Presentation.SlideIdList.Count();

                // Проходим по всем слайдам по из порядку
                slideIds = presPart.Presentation.SlideIdList.Cast<SlideId>().ToList();
                for (int i = 1; i < slideIds.Count; i++) // пропускаем титульный слайд
                {
                    if (slideIds[i].RelationshipId == tocRelId) continue; // пропускаем TOC

                    var slidePart = (SlidePart)presPart.GetPartById(slideIds[i].RelationshipId);
                    AddBacklinkToSlide(slidePart, tocPart, presPart, i + 1, totalSlides);
                }

                presPart.Presentation.Save();
            }
        }

        private void AddBacklinkToSlide(
            SlidePart slidePart,
            SlidePart tocPart,
            PresentationPart presPart,
            int currentSlideNumber,
            int totalSlides)
        {
            var shapeTree = slidePart.Slide.CommonSlideData.ShapeTree;

            // Генерация уникального ID
            uint maxId = 1;
            foreach (var el in shapeTree.Elements())
            {
                if (el is DocumentFormat.OpenXml.Presentation.Shape s && s.NonVisualShapeProperties?.NonVisualDrawingProperties?.Id != null)
                    maxId = Math.Max(maxId, s.NonVisualShapeProperties.NonVisualDrawingProperties.Id.Value + 1);
                else if (el is Picture p && p.NonVisualPictureProperties?.NonVisualDrawingProperties?.Id != null)
                    maxId = Math.Max(maxId, p.NonVisualPictureProperties.NonVisualDrawingProperties.Id.Value + 1);
            }

            // Размер слайда
            long slideWidth = presPart.Presentation.SlideSize.Cx;
            long slideHeight = presPart.Presentation.SlideSize.Cy;

            long margin = 300000;
            long textWidth = 800000;
            long textHeight = 200000;
            long x = slideWidth - textWidth - margin;
            long y = slideHeight - textHeight - margin;

            string backlinkText = $"{currentSlideNumber}/{totalSlides}";

            // добавляем TOC-слайд как часть текущего слайда
            slidePart.AddPart(tocPart); // создаёт локальное отношение
            string localTocRelId = slidePart.GetIdOfPart(tocPart); // локальный ID

            var backlinkShape = new DocumentFormat.OpenXml.Presentation.Shape(
                new NonVisualShapeProperties(
                    new NonVisualDrawingProperties { Id = maxId, Name = "Backlink to TOC" },
                    new NonVisualShapeDrawingProperties(),
                    new ApplicationNonVisualDrawingProperties()
                ),
                new ShapeProperties(
                    new A.Transform2D(
                        new A.Offset { X = x, Y = y },
                        new A.Extents { Cx = textWidth, Cy = textHeight }
                    ),
                    new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle },
                    new A.NoFill(),
                    new A.Outline { Width = 0 }
                ),
                new TextBody(
                    new A.BodyProperties(),
                    new A.ListStyle(),
                    new A.Paragraph(
                        new A.Run(
                            new A.RunProperties(
                                new A.SolidFill(new A.RgbColorModelHex { Val = "0000FF" }),
                                new A.Underline(),
                                new A.HyperlinkOnClick
                                {
                                    Id = localTocRelId,
                                    Action = "ppaction://hlinksldjump"
                                })
                            {
                                Language = "ru-RU"
                            },
                            new A.Text(backlinkText)
                        )
                    )
                )
            );

            shapeTree.Append(backlinkShape);
        }


        private string GetOutputPath(string inputPath)
        {
            string dir = Path.GetDirectoryName(inputPath);
            string name = Path.GetFileNameWithoutExtension(inputPath);
            string ext = Path.GetExtension(inputPath);
            int i = 1;
            string path;
            do { path = Path.Combine(dir, name + (i == 1 ? "_TOC" : $"_TOC({i})") + ext); i++; }
            while (File.Exists(path));
            return path;
        }

        private static void TryDeleteFolder(string path)
        {
            try { if (Directory.Exists(path)) Directory.Delete(path, true); } catch { }
        }

        public void Dispose()
        {
            if (!_disposed) { _disposed = true; GC.SuppressFinalize(this); }
        }



        //public string CreateTableOfContents(string inputPath, List<Models.SlideItem> slides, int columns, int margin)
        //{
        //    Application pptApp = null;
        //    Presentation pres = null;
        //    string outputPath = GetOutputPath(inputPath);

        //    try
        //    {
        //        File.Copy(inputPath, outputPath, true);

        //        pptApp = new Application { DisplayAlerts = PpAlertLevel.ppAlertsNone };
        //        pres = pptApp.Presentations.Open(outputPath, MsoTriState.msoFalse, MsoTriState.msoFalse, MsoTriState.msoFalse);

        //        // Добавляем слайд оглавления
        //        Slide firstSlide = pres.Slides[1];
        //        Slide tocSlide = pres.Slides.Add(pres.Slides.Count + 1, firstSlide.Layout);
        //        tocSlide.Name = "Оглавление";
        //        tocSlide.Design = firstSlide.Design;
        //        tocSlide.ColorScheme = firstSlide.ColorScheme;

        //        RemoveDefaultShapes(tocSlide);
        //        AddTitle(tocSlide, "ОГЛАВЛЕНИЕ");

        //        // Создаем временную папку для миниатюр
        //        string tempDir = Path.Combine(Path.GetTempPath(), "toc_thumbs_" + Guid.NewGuid());
        //        Directory.CreateDirectory(tempDir);

        //        try
        //        {
        //            float slideWidth = pres.PageSetup.SlideWidth;
        //            float slideHeight = pres.PageSetup.SlideHeight;

        //            var layout = LayoutCalculatorService.CalculateOptimalLayout(slides.Count, margin, columns, slideWidth, slideHeight);
        //            AddThumbnails(pres, tocSlide, slides, layout.Columns, margin, tempDir, layout);
        //        }
        //        finally
        //        {
        //            TryDeleteFolder(tempDir);
        //        }

        //        pres.Save();
        //    }
        //    finally
        //    {
        //        if (pres != null) { pres.Close(); Marshal.ReleaseComObject(pres); }
        //        if (pptApp != null) { pptApp.Quit(); Marshal.ReleaseComObject(pptApp); }
        //    }

        //    ApplyOpenXmlHyperlinks(outputPath);
        //    return outputPath;
        //}

        //private void AddThumbnails(Presentation pres, Slide tocSlide, List<Models.SlideItem> slides, int columns, int margin, string tempDir, (int Columns, float ThumbWidth, float ThumbHeight, float RowHeight) layout)
        //{
        //    float yStart = LayoutConstants.TitleHeight;
        //    float captionHeight = LayoutConstants.CaptionHeight;

        //    for (int i = 0; i < slides.Count; i++)
        //    {
        //        int row = i / columns;
        //        int col = i % columns;

        //        float x = margin + col * (layout.ThumbWidth + margin);
        //        float y = yStart + row * layout.RowHeight;

        //        AddThumbnail(pres, tocSlide, slides[i], x, y, layout.ThumbWidth, layout.ThumbHeight, captionHeight, tempDir);
        //    }
        //}

        //private void AddThumbnail(Presentation pres, Slide tocSlide, Models.SlideItem slideItem, float x, float y, float width, float height, float captionHeight, string tempDir)
        //{
        //    string path = Path.Combine(tempDir, $"slide_{slideItem.Number}.png");
        //    pres.Slides[slideItem.Number].Export(path, "PNG", (int)(width * 4), (int)(height * 4));

        //    string guid = $"thumb_{slideItem.Number}";
        //    Shape pic = tocSlide.Shapes.AddPicture(path, MsoTriState.msoFalse, MsoTriState.msoTrue, x, y, width, height);
        //    pic.Name = guid;

        //    Shape caption = tocSlide.Shapes.AddTextbox(
        //        MsoTextOrientation.msoTextOrientationHorizontal,
        //        x, y + height, width, captionHeight);
        //    caption.TextFrame.TextRange.Text = $"Слайд {slideItem.Number}";
        //    caption.TextFrame.TextRange.Font.Size = 12;
        //    caption.TextFrame.TextRange.Font.Bold = MsoTriState.msoTrue;
        //    caption.TextFrame.TextRange.ParagraphFormat.Alignment = PpParagraphAlignment.ppAlignCenter;

        //    // Make caption background semi-transparent
        //    caption.Fill.ForeColor.RGB = 0xFFFFFF;
        //    caption.Fill.Transparency = 0.7f;

        //    Shape frame = tocSlide.Shapes.AddShape(MsoAutoShapeType.msoShapeRectangle, x, y, width, height);
        //    frame.Fill.Visible = MsoTriState.msoFalse;
        //    frame.Line.ForeColor.RGB = 0x808080;
        //    frame.Line.Weight = 1.5f;

        //    _pendingLinks.Add(new TocLinkInfo { ShapeGuid = guid, SlideNumber = slideItem.Number });
        //}

        //private void AddTitle(Slide slide, string text)
        //{
        //    float width = slide.Parent.PageSetup.SlideWidth - 100;
        //    Shape title = slide.Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal, 50, 30, width, 60);
        //    title.TextFrame.TextRange.Text = text;
        //    title.TextFrame.TextRange.Font.Size = 36;
        //    title.TextFrame.TextRange.Font.Bold = MsoTriState.msoTrue;
        //    title.TextFrame.TextRange.ParagraphFormat.Alignment = PpParagraphAlignment.ppAlignCenter;
        //}

        //private void RemoveDefaultShapes(Slide slide)
        //{
        //    foreach (Shape s in slide.Shapes.Cast<Shape>().ToList())
        //        if (s.Type == MsoShapeType.msoPlaceholder)
        //            s.Delete();
        //}

        //private void ApplyOpenXmlHyperlinks(string pptxPath)
        //{
        //    using (PresentationDocument doc = PresentationDocument.Open(pptxPath, true))
        //    {
        //        var presPart = doc.PresentationPart;
        //        if (presPart == null) return;

        //        var slideParts = presPart.SlideParts.ToList();
        //        if (!slideParts.Any()) return;

        //        // TOC slide - последний
        //        var tocSlidePart = slideParts.Last();
        //        var tocSlide = tocSlidePart.Slide;

        //        foreach (var pic in tocSlide.Descendants<DocumentFormat.OpenXml.Presentation.Picture>())
        //        {
        //            var shapeName = pic.NonVisualPictureProperties?.NonVisualDrawingProperties?.Name?.Value;

        //            if (shapeName?.StartsWith("thumb_") == true)
        //            {
        //                if (int.TryParse(shapeName.Substring(6), out int slideNumber))
        //                {
        //                    if (slideNumber >= 1 && slideNumber <= slideParts.Count)
        //                    {
        //                        var targetSlidePart = slideParts[slideNumber - 1];

        //                        tocSlidePart.AddPart(targetSlidePart);

        //                        string relId = tocSlidePart.GetIdOfPart(targetSlidePart);

        //                        pic.NonVisualPictureProperties
        //                           .NonVisualDrawingProperties
        //                           .HyperlinkOnClick = new A.HyperlinkOnClick
        //                           {
        //                               Id = relId,
        //                               Action = "ppaction://hlinksldjump"
        //                           };

        //                    }
        //                }
        //            }
        //        }

        //        tocSlide.Save();
        //    }
        //}
    }
}
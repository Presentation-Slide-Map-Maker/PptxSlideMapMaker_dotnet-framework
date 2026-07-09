using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using TocBuilder_dotnet_framework.Models;
using A = DocumentFormat.OpenXml.Drawing;

namespace TocBuilder_dotnet_framework.Services
{
    public static class TocConstants
    {
        public const string CreatorName = "TOC Builder";
        public const string NotesText = "Made with TOC Builder";
        public const string DefaultLanguage = "ru-RU";
        public const string SlideCaptionPrefix = "Слайд ";
        public const string BacklinkName = "Backlink to TOC";
    }

    public class TocGeneratorService
    {
        private readonly SlideExportService _slideExportService;

        public TocGeneratorService()
        {
            _slideExportService = new SlideExportService();
        }

        public string CreateTableOfContents(
            string inputPath,
            List<Models.SlideItem> slides,
            int columns,
            int margin,
            int backgroundSlideIndex,
            string selectedFont = null,
            int selectedFontSize = 12)
        {
            string outputPath = GetOutputPath(inputPath);

            // Copy input presentation file
            File.Copy(inputPath, outputPath, true);

            // Create temporary folder for exported slide thumbnails
            string tempDir = Path.Combine(
                Path.GetTempPath(),
                "toc_thumbs_" + Guid.NewGuid().ToString("N"));

            Directory.CreateDirectory(tempDir);

            try
            {
                var slideNumbers = slides.Select(s => s.Number).ToList();

                // Export slides to PNG
                _slideExportService.ExportSlidesToPng(outputPath, tempDir, slideNumbers, 1600, 900);

                // Build Table of Contents slide
                BuildTocSlideOpenXml(
                    outputPath,
                    slides,
                    columns,
                    margin,
                    tempDir,
                    backgroundSlideIndex,
                    selectedFont,
                    selectedFontSize);

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
            int backgroundSlideIndex,
            string selectedFont = null,
            int selectedFontSize = 12)
        {
            using (var doc = PresentationDocument.Open(pptxPath, true))
            {
                // Указание авторства
                var props = doc.PackageProperties;
                props.Creator = TocConstants.CreatorName;
                props.LastModifiedBy = TocConstants.CreatorName;

                var presPart = doc.PresentationPart;
                var slideIds = presPart.Presentation.SlideIdList.Elements<SlideId>().ToList();

                // Валидация slide index
                if (backgroundSlideIndex < 0 || backgroundSlideIndex >= slideIds.Count)
                    throw new ArgumentOutOfRangeException(nameof(backgroundSlideIndex));

                var backgroundSlideId = slideIds[backgroundSlideIndex];
                var backgroundSlidePart = (SlidePart)presPart.GetPartById(backgroundSlideId.RelationshipId);
                var layoutPart = backgroundSlidePart.SlideLayoutPart;

                // TOC Slide Part
                var tocPart = presPart.AddNewPart<SlidePart>();
                tocPart.Slide = new DocumentFormat.OpenXml.Presentation.Slide(
                    new CommonSlideData(new ShapeTree()),
                    new ColorMapOverride(new A.MasterColorMapping()));

                // Notes Slide Part
                var notesPart = tocPart.AddNewPart<NotesSlidePart>();
                CreateNotesSlide(notesPart, TocConstants.NotesText);
                notesPart.AddPart(presPart.NotesMasterPart); // Bind to NotesMaster

                tocPart.AddPart(layoutPart);
                tocPart.Slide.Save();

                // Добавление TOC Slide в Presentation Slide List
                presPart.Presentation.SlideIdList.Append(
                    new SlideId
                    {
                        Id = slideIds.Max(s => s.Id.Value) + 1,
                        RelationshipId = presPart.GetIdOfPart(tocPart)
                    });

                var shapeTree = tocPart.Slide.CommonSlideData.ShapeTree;

                float slideWidthPx = presPart.Presentation.SlideSize.Cx / (float)OpenXmlUnits.EmuPerPixelAt96Dpi;
                float slideHeightPx = presPart.Presentation.SlideSize.Cy / (float)OpenXmlUnits.EmuPerPixelAt96Dpi;

                var layout = LayoutCalculatorService.CalculateOptimalLayout(
                    slides.Count,
                    margin,
                    columns,
                    slideWidthPx,
                    slideHeightPx
                );

                // Calculate total width of the grid area
                float gridWidth = layout.Columns * layout.ThumbWidth + (layout.Columns - 1) * margin;
                // Center the grid by sharing remaining horizontal space equally
                float xStart = (slideWidthPx - gridWidth) / 2f;

                for (int i = 0; i < slides.Count; i++)
                {
                    int row = i / layout.Columns;
                    int col = i % layout.Columns;

                    long x = OpenXmlUnits.PixelsToEmu(xStart + col * (layout.ThumbWidth + margin));
                    long y = OpenXmlUnits.PixelsToEmu(LayoutConstants.TitleHeight + row * layout.RowHeight);

                    long thumbW = OpenXmlUnits.PixelsToEmu(layout.ThumbWidth);
                    long thumbH = OpenXmlUnits.PixelsToEmu(layout.ThumbHeight);
                    long captionH = OpenXmlUnits.PixelsToEmu(LayoutConstants.CaptionHeight);

                    var targetSlidePart = (SlidePart)presPart.GetPartById(slideIds[slides[i].Number - 1].RelationshipId);

                    tocPart.AddPart(targetSlidePart);
                    string relId = tocPart.GetIdOfPart(targetSlidePart);

                    var imgPart = tocPart.AddImagePart(ImagePartType.Png);
                    using (var fs = File.OpenRead(Path.Combine(tempDir, $"slide_{slides[i].Number}.png")))
                        imgPart.FeedData(fs);

                    string imgRelId = tocPart.GetIdOfPart(imgPart);

                    // 1. Frame Shape
                    AddThumbnailFrame(shapeTree, i, x, y, thumbW, thumbH, slides[i].Number);

                    // 2. Picture Shape
                    AddThumbnailPicture(shapeTree, i, x, y, thumbW, thumbH, imgRelId, relId, slides[i].Number);

                    // 3. Caption Shape
                    long captionY = y + thumbH;
                    AddThumbnailCaption(shapeTree, i, x, captionY, thumbW, captionH, relId, slides[i].Number, selectedFont, selectedFontSize);
                }

                // Add backlinks from other slides back to TOC
                string tocRelId = presPart.GetIdOfPart(tocPart);
                int totalSlides = presPart.Presentation.SlideIdList.Count();

                slideIds = presPart.Presentation.SlideIdList.Cast<SlideId>().ToList();
                for (int i = 1; i < slideIds.Count; i++) // Пропускаем title slide
                {
                    if (slideIds[i].RelationshipId == tocRelId) continue; // пропускаем TOC slide

                    var slidePart = (SlidePart)presPart.GetPartById(slideIds[i].RelationshipId);
                    AddBacklinkToSlide(slidePart, tocPart, presPart, i + 1, totalSlides, selectedFont, selectedFontSize);
                }

                presPart.Presentation.Save();
            }
        }

        private void CreateNotesSlide(NotesSlidePart notesPart, string notesText)
        {
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
                                        new A.Text(notesText)
                                    )
                                )
                            )
                        )
                    )
                )
            );
        }

        private void AddThumbnailFrame(
            ShapeTree shapeTree,
            int index,
            long x,
            long y,
            long width,
            long height,
            int slideNumber)
        {
            shapeTree.Append(
                new DocumentFormat.OpenXml.Presentation.Shape(
                    new NonVisualShapeProperties(
                        new NonVisualDrawingProperties
                        {
                            Id = (uint)(3000 + index),
                            Name = $"Frame {slideNumber}"
                        },
                        new NonVisualShapeDrawingProperties(),
                        new ApplicationNonVisualDrawingProperties()
                    ),
                    new ShapeProperties(
                        new A.Transform2D(
                            new A.Offset { X = x, Y = y },
                            new A.Extents { Cx = width, Cy = height }),
                        new A.PresetGeometry(new A.AdjustValueList())
                        { Preset = A.ShapeTypeValues.Rectangle },
                        new A.Outline(
                            new A.SolidFill(
                                new A.RgbColorModelHex { Val = "808080" }
                            )
                        ),
                        new A.NoFill()
                    )
                )
            );
        }

        private void AddThumbnailPicture(
            ShapeTree shapeTree,
            int index,
            long x,
            long y,
            long width,
            long height,
            string imgRelId,
            string targetSlideRelId,
            int slideNumber)
        {
            shapeTree.Append(
                new Picture(
                    new NonVisualPictureProperties(
                        new NonVisualDrawingProperties
                        {
                            Id = (uint)(2000 + index),
                            Name = $"Slide {slideNumber}",
                            HyperlinkOnClick = new A.HyperlinkOnClick
                            {
                                Id = targetSlideRelId,
                                Action = "ppaction://hlinksldjump"
                            }
                        },
                        new NonVisualPictureDrawingProperties(),
                        new ApplicationNonVisualDrawingProperties()
                    ),
                    new BlipFill(
                        new A.Blip { Embed = imgRelId },
                        new A.Stretch(new A.FillRectangle())
                    ),
                    new ShapeProperties(
                        new A.Transform2D(
                            new A.Offset { X = x, Y = y },
                            new A.Extents { Cx = width, Cy = height }),
                        new A.PresetGeometry(new A.AdjustValueList())
                        { Preset = A.ShapeTypeValues.Rectangle })
                ));
        }

        private void AddThumbnailCaption(
            ShapeTree shapeTree,
            int index,
            long x,
            long y,
            long width,
            long height,
            string targetSlideRelId,
            int slideNumber,
            string selectedFont = null,
            int selectedFontSize = 12)
        {
            var runPr = new A.RunProperties(
                new A.HyperlinkOnClick
                {
                    Id = targetSlideRelId,
                    Action = "ppaction://hlinksldjump"
                })
            {
                Language = TocConstants.DefaultLanguage,
                FontSize = selectedFontSize * 100
            };
            if (!string.IsNullOrEmpty(selectedFont))
            {
                runPr.Append(new A.LatinFont { Typeface = selectedFont });
                runPr.Append(new A.ComplexScriptFont { Typeface = selectedFont });
                runPr.Append(new A.EastAsianFont { Typeface = selectedFont });
            }

            shapeTree.Append(
                new DocumentFormat.OpenXml.Presentation.Shape(
                    new NonVisualShapeProperties(
                        new NonVisualDrawingProperties
                        {
                            Id = (uint)(4000 + index),
                            Name = $"Caption {slideNumber}"
                        },
                        new NonVisualShapeDrawingProperties(),
                        new ApplicationNonVisualDrawingProperties()
                    ),
                    new ShapeProperties(
                        new A.Transform2D(
                            new A.Offset { X = x, Y = y },
                            new A.Extents { Cx = width, Cy = height }),
                        new A.PresetGeometry(new A.AdjustValueList())
                        { Preset = A.ShapeTypeValues.Rectangle }
                    ),
                    new TextBody(
                        new A.BodyProperties(),
                        new A.ListStyle(),
                        new A.Paragraph(
                            new A.Run(
                                runPr,
                                new A.Text($"{TocConstants.SlideCaptionPrefix}{slideNumber}")
                            )
                        )
                    )
                )
            );
        }

        private void AddBacklinkToSlide(
            SlidePart slidePart,
            SlidePart tocPart,
            PresentationPart presPart,
            int currentSlideNumber,
            int totalSlides,
            string selectedFont = null,
            int selectedFontSize = 12)
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

            // Backlink dimensions in Points converted to EMUs
            const double backlinkMarginPt = 24.0;
            const double backlinkWidthPt = 64.0;
            const double backlinkHeightPt = 16.0;

            long margin = OpenXmlUnits.PointsToEmu(backlinkMarginPt);
            long textWidth = OpenXmlUnits.PointsToEmu(backlinkWidthPt);
            long textHeight = OpenXmlUnits.PointsToEmu(backlinkHeightPt);
            long x = slideWidth - textWidth - margin;
            long y = slideHeight - textHeight - margin;

            string backlinkText = $"{currentSlideNumber}/{totalSlides}";

            // добавляем TOC-слайд как часть текущего слайда
            slidePart.AddPart(tocPart); // создаёт локальное отношение
            string localTocRelId = slidePart.GetIdOfPart(tocPart); // локальный ID

            var runPr = new A.RunProperties(
                new A.SolidFill(new A.RgbColorModelHex { Val = "0000FF" }),
                new A.Underline(),
                new A.HyperlinkOnClick
                {
                    Id = localTocRelId,
                    Action = "ppaction://hlinksldjump"
                })
            {
                Language = TocConstants.DefaultLanguage,
                FontSize = selectedFontSize * 100
            };
            if (!string.IsNullOrEmpty(selectedFont))
            {
                runPr.Append(new A.LatinFont { Typeface = selectedFont });
                runPr.Append(new A.ComplexScriptFont { Typeface = selectedFont });
                runPr.Append(new A.EastAsianFont { Typeface = selectedFont });
            }

            var backlinkShape = new DocumentFormat.OpenXml.Presentation.Shape(
                new NonVisualShapeProperties(
                    new NonVisualDrawingProperties { Id = maxId, Name = TocConstants.BacklinkName },
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
                            runPr,
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
    }
}
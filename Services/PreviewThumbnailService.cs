using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using TocBuilder_dotnet_framework.Models;

namespace TocBuilder_dotnet_framework.Services
{
    public class PreviewThumbnailService
    {
        private readonly SlideExportService _slideExportService;

        public PreviewThumbnailService()
        {
            _slideExportService = new SlideExportService();
        }

        public List<Models.SlideItem> GetSlides(string filePath)
        {
            var slides = new List<Models.SlideItem>();
            string tempDir = Path.Combine(Path.GetTempPath(), $"ppt_thumbs_{Guid.NewGuid()}");
            Directory.CreateDirectory(tempDir);

            try
            {
                int slideCount = 0;
                float slideWidth = LayoutConstants.DefaultSlideWidth;
                float slideHeight = LayoutConstants.DefaultSlideHeight;

                using (var doc = PresentationDocument.Open(filePath, isEditable: false))
                {
                    var presPart = doc.PresentationPart;
                    if (presPart?.Presentation?.SlideIdList != null)
                    {
                        slideCount = presPart.Presentation.SlideIdList.Count();
                    }

                    var slideSize = presPart?.Presentation?.SlideSize;
                    if (slideSize != null && slideSize.Cx.HasValue && slideSize.Cy.HasValue)
                    {
                        slideWidth = (float)(slideSize.Cx.Value / OpenXmlUnits.EmuPerPixelAt96Dpi);
                        slideHeight = (float)(slideSize.Cy.Value / OpenXmlUnits.EmuPerPixelAt96Dpi);
                    }
                }

                if (slideCount == 0) return slides;

                float aspect = slideWidth / slideHeight;
                int previewWidth = 320;
                int previewHeight = (int)(previewWidth / aspect);

                var slideNumbers = Enumerable.Range(1, slideCount).ToList();
                
                // Ёкспорт через SlideExportService
                _slideExportService.ExportSlidesToPng(filePath, tempDir, slideNumbers, previewWidth * 2, previewHeight * 2);

                for (int i = 1; i <= slideCount; i++)
                {
                    string thumbPath = Path.Combine(tempDir, $"slide_{i}.png");
                    
                    if (!File.Exists(thumbPath))
                    {
                        CreateFallbackThumbnail(thumbPath, i, previewWidth, previewHeight);
                    }

                    byte[] thumbnailBytes = File.ReadAllBytes(thumbPath);

                    slides.Add(new Models.SlideItem
                    {
                        Number = i,
                        Thumbnail = thumbnailBytes,
                        IsSelected = true
                    });
                }
            }
            finally
            {
                TryDeleteFolder(tempDir);
            }

            return slides;
        }

        private void CreateFallbackThumbnail(string filePath, int slideNumber, int width, int height)
        {
            using (Bitmap bmp = new Bitmap(width, height))
            using (Graphics g = Graphics.FromImage(bmp))
            {
                g.Clear(Color.LightGray);
                g.DrawString($"Slide {slideNumber}",
                    new Font("Arial", 20),
                    Brushes.Black,
                    new PointF(10, 10));
                bmp.Save(filePath, ImageFormat.Png);
            }
        }

        public (float Width, float Height) GetSlideDimensions(string pptxPath)
        {
            try
            {
                using (var doc = PresentationDocument.Open(pptxPath, isEditable: false))
                {
                    var presentationPart = doc.PresentationPart;
                    var presentation = presentationPart?.Presentation;

                    var slideSize = presentation?.SlideSize;
                    if (slideSize != null)
                    {
                        long? cx = slideSize.Cx;
                        long? cy = slideSize.Cy;

                        if (cx.HasValue && cy.HasValue)
                        {
                            const double EMU_PER_PIXEL = 9525.0; // 1 pixel = 9525 EMU at 96 DPI
                            float widthPx = (float)(cx.Value / EMU_PER_PIXEL);
                            float heightPx = (float)(cy.Value / EMU_PER_PIXEL);
                            return (widthPx, heightPx);
                        }
                    }
                }
            }
            catch
            {
                // Fallback dimensions logic handled below
            }

            return (LayoutConstants.DefaultSlideWidth, LayoutConstants.DefaultSlideHeight);
        }

        private static void TryDeleteFolder(string path)
        {
            try { if (Directory.Exists(path)) Directory.Delete(path, true); } catch { }
        }
    }
}

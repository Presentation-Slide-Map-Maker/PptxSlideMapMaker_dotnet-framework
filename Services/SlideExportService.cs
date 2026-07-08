using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Application = Microsoft.Office.Interop.PowerPoint.Application;

namespace TocBuilder_dotnet_framework.Services
{
    public class SlideExportService
    {
        public void ExportSlidesToPng(string presentationPath, string outputDirectory, IEnumerable<int> slideNumbers, int width, int height)
        {
            Application pptApp = null;
            Presentations presentations = null;
            Presentation pres = null;
            Slides slides = null;

            try
            {
                pptApp = new Application { DisplayAlerts = PpAlertLevel.ppAlertsNone };
                presentations = pptApp.Presentations;
                pres = presentations.Open(presentationPath, MsoTriState.msoTrue, MsoTriState.msoFalse, MsoTriState.msoFalse);
                slides = pres.Slides;

                foreach (int number in slideNumbers)
                {
                    Slide slide = null;
                    try
                    {
                        slide = slides[number];
                        string path = Path.Combine(outputDirectory, $"slide_{number}.png");
                        slide.Export(path, "PNG", width, height);
                    }
                    catch (Exception)
                    {
                        
                    }
                    finally
                    {
                        if (slide != null) Marshal.ReleaseComObject(slide);
                    }
                }
            }
            finally
            {
                if (pres != null) { pres.Close(); Marshal.ReleaseComObject(pres); }
                if (presentations != null) Marshal.ReleaseComObject(presentations);
                if (slides != null) Marshal.ReleaseComObject(slides);
                if (pptApp != null) { pptApp.Quit(); Marshal.ReleaseComObject(pptApp); }

                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using TocBuilder_dotnet_framework.Models;

namespace TocBuilder_dotnet_framework.Services
{
    public static class LayoutConstants
    {
        public const float TitleHeight = 120f;
        public const float CaptionHeight = 20f;
        public const float BottomMargin = 40f;
        public const float MinThumbWidth = 80f;
        public const float MinThumbHeight = 60f;

        public const float DefaultSlideWidth = 800f;   // Ширина слайда по умолчанию (13.333 in × 72 = 960)
        public const float DefaultSlideHeight = 450f;  // Высота слайда по умолчанию (7.5 in × 72 = 540)
    }

    public static class LayoutCalculatorService
    {
        public static (int Columns, float ThumbWidth, float ThumbHeight, float RowHeight) CalculateOptimalLayout(
            int slideCount,
            float margin,
            int desiredColumns = -1,
            float slideWidth = LayoutConstants.DefaultSlideWidth,
            float slideHeight = LayoutConstants.DefaultSlideHeight)
        {
            float slideAspectRatio = slideWidth / slideHeight;
            const float titleHeight = LayoutConstants.TitleHeight;
            const float bottomMargin = LayoutConstants.BottomMargin;
            const float captionHeight = LayoutConstants.CaptionHeight;
            const float minThumbWidth = LayoutConstants.MinThumbWidth;
            const float minThumbHeight = LayoutConstants.MinThumbHeight;

            // Доступная рабочая область по высоте и ширине
            float availableHeight = slideHeight - titleHeight - bottomMargin;
            float availableWidth = slideWidth - (margin * 2);

            // Локальная функция для расчета размеров миниатюры при заданном числе колонок
            (bool Fits, float ThumbW, float ThumbH, float RowH) CalculateSizeForColumns(int cols)
            {
                if (cols <= 0) return (false, 0, 0, 0);

                int rows = (int)Math.Ceiling((double)slideCount / cols);

                // Исходный расчет ширины и высоты миниатюры на основе ограничения по ширине
                float thumbW = (availableWidth - margin * (cols - 1)) / cols;
                float thumbH = thumbW / slideAspectRatio;

                float totalH = rows * (thumbH + margin + captionHeight);

                // Если высота сетки превышает доступную — пропорционально сжимаем под высоту
                if (totalH > availableHeight)
                {
                    float maxThumbH = (availableHeight - rows * (captionHeight + margin)) / rows;
                    if (maxThumbH > 0)
                    {
                        thumbH = maxThumbH;
                        thumbW = thumbH * slideAspectRatio;
                        totalH = rows * (thumbH + margin + captionHeight); // Пересчитываем общую высоту
                    }
                }

                // Проверяем, укладывается ли сетка в доступную высоту и соблюдаются ли минимальные размеры
                bool fits = totalH <= availableHeight && thumbW >= minThumbWidth && thumbH >= minThumbHeight;
                float rowH = thumbH + margin + captionHeight;

                return (fits, thumbW, thumbH, rowH);
            }

            // 1. Если задано желаемое число колонок вручную — используем его
            if (desiredColumns > 0)
            {
                var result = CalculateSizeForColumns(desiredColumns);
                if (result.Fits)
                    return (desiredColumns, result.ThumbW, result.ThumbH, result.RowH);
            }

            // 2. В автоматическом режиме перебираем от 1 до 6 колонок и ищем ту разметку, где миниатюры крупнее
            int bestCols = 1;
            float bestW = 0;
            float bestH = 0;
            float bestRH = 0;

            int maxColsToCheck = Math.Min(6, slideCount);

            for (int c = 1; c <= maxColsToCheck; c++)
            {
                var (fits, w, h, rh) = CalculateSizeForColumns(c);
                if (!fits) continue;

                // Выбираем разметку, которая максимизирует ширину миниатюры
                if (w > bestW)
                {
                    bestW = w;
                    bestH = h;
                    bestCols = c;
                    bestRH = rh;
                }
            }

            // 3. Крайний fallback: ни одна разметка не подошла
            if (bestW == 0)
            {
                var fallback = CalculateSizeForColumns(1);
                if (fallback.Fits)
                    return (1, fallback.ThumbW, fallback.ThumbH, fallback.RowH);

                // Жестко принудительно используем минимальные размеры
                float w = minThumbWidth;
                float h = Math.Max(minThumbHeight, minThumbWidth / slideAspectRatio);
                float rh = h + margin + captionHeight;
                return (Math.Min(4, slideCount), w, h, rh);
            }

            return (bestCols, bestW, bestH, bestRH);
        }

        // Генерация элементов превью для Canvas
        public static List<PreviewItem> GeneratePreviewItems(
            List<SlideItem> selectedSlides,
            int columns,
            int margin,
            float slideWidth = LayoutConstants.DefaultSlideWidth,
            float slideHeight = LayoutConstants.DefaultSlideHeight)
        {
            var previewItems = new List<PreviewItem>();

            if (selectedSlides == null || !selectedSlides.Any()) return previewItems;

            var layoutInfo = CalculateOptimalLayout(
                selectedSlides.Count,
                margin,
                columns,
                slideWidth,
                slideHeight);

            float thumbWidth = layoutInfo.ThumbWidth;
            float thumbHeight = layoutInfo.ThumbHeight;
            float rowHeight = layoutInfo.RowHeight;
            int actualCols = layoutInfo.Columns;

            // Расчет общей ширины сетки миниатюр
            float gridWidth = actualCols * thumbWidth + (actualCols - 1) * margin;
            // Центрирование сетки по горизонтали
            float xStart = (slideWidth - gridWidth) / 2f;

            const float titleHeight = LayoutConstants.TitleHeight;
            const float yStart = titleHeight;

            for (int i = 0; i < selectedSlides.Count; i++)
            {
                int row = i / actualCols;
                int col = i % actualCols;

                float x = xStart + col * (thumbWidth + margin);
                float y = yStart + row * rowHeight;

                previewItems.Add(new PreviewItem
                {
                    X = x,
                    Y = y,
                    Width = thumbWidth,
                    Height = thumbHeight,
                    Caption = $"Слайд {selectedSlides[i].Number}",
                    Thumbnail = selectedSlides[i].Thumbnail
                });
            }

            return previewItems;
        }
    }
}
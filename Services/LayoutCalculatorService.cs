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

        public const float DefaultSlideWidth = 800f;   // 13.333 in × 72 = 960 (16:9)
        public const float DefaultSlideHeight = 450f;  // 7.5 in × 72 = 540 (16:9)
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

            // Доступная высота на мирниатюры и подписи
            float availableHeight = slideHeight - titleHeight - bottomMargin;
            float availableWidth = slideWidth - (margin * 2);

            // Вспомогательная функция для проверки конфигурации
            (bool Fits, float ThumbW, float ThumbH, float RowH) TryColumns(int cols)
            {
                // Ширина каждой миниатюры при кол-ве колонок = cols
                float thumbW = (availableWidth - margin * (cols - 1)) / cols;
                float thumbH = thumbW / slideAspectRatio;

                int rows = (int)Math.Ceiling((double)slideCount / cols);
                float rowH = thumbH + margin + captionHeight;

                float totalH = rows * rowH; // Общая высота сетки
                bool fits = totalH <= availableHeight && thumbW >= minThumbWidth && thumbH >= minThumbHeight;

                // Если не помещается по высоте — попробуем сжать
                if (!fits && totalH > availableHeight)
                {
                    // Доступная высота на миниатюры и подписи
                    float maxThumbH = (availableHeight - rows * (captionHeight + margin)) / rows;
                    if (maxThumbH > 0)
                    {
                        thumbH = Math.Max(minThumbHeight, maxThumbH);
                        thumbW = thumbH * slideAspectRatio; // сохраняем соотношение

                        // Проверим, умещается ли ширина
                        if (thumbW > (availableWidth - margin * (cols - 1)) / cols)
                        {
                            thumbW = (availableWidth - margin * (cols - 1)) / cols;
                            thumbH = thumbW / slideAspectRatio;
                        }

                        rowH = thumbH + margin + captionHeight;
                        totalH = rows * rowH;
                        fits = totalH <= availableHeight && thumbW >= minThumbWidth && thumbH >= minThumbHeight;
                    }
                }

                return (fits, thumbW, thumbH, rowH);
            }

            // 1. Если задано желаемое количество колонок — попробуем его сначала
            if (desiredColumns > 0)
            {
                var result = TryColumns(desiredColumns);
                if (result.Fits)
                    return (desiredColumns, result.ThumbW, result.ThumbH, result.RowH);
            }

            // 2. Иначе — перебор количества колонок, ищем наибольшую площадь миниатюры
            (int bestCols, float bestW, float bestH, float bestRH) = (1, 0, 0, 0);
            float bestScore = 0;

            for (int r = 1; r <= slideCount; r++)
            {
                int c = (int)Math.Ceiling((double)slideCount / r);
                var (fits, w, h, rh) = TryColumns(c);
                if (!fits) continue;

                float area = w * h;
                float spaceUsed = (r * rh) / availableHeight;
                int emptyCells = r * c - slideCount;
                float symmetry = 1.0f / (1.0f + Math.Abs(r - c));

                // Эвристика
                float score = area * spaceUsed * symmetry * (1.0f / (emptyCells + 1));

                if (score > bestScore)
                {
                    bestScore = score;
                    (bestCols, bestW, bestH, bestRH) = (c, w, h, rh);
                }
            }

            // 3. Если ничего не подошло — используем fallback (1 колонка, принудительное сжатие)
            if (bestScore == 0)
            {
                var fallback = TryColumns(1);
                if (fallback.Fits)
                    return (1, fallback.ThumbW, fallback.ThumbH, fallback.RowH);

                // Крайний fallback: жёстко используем минимальные размеры
                float w = minThumbWidth;
                float h = Math.Max(minThumbHeight, minThumbWidth / slideAspectRatio); // сохраняем пропорции
                float rh = h + margin + captionHeight;
                return (Math.Min(4, slideCount), w, h, rh);
            }

            return (bestCols, bestW, bestH, bestRH);
        }

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

            const float titleHeight = LayoutConstants.TitleHeight;
            const float yStart = titleHeight;

            for (int i = 0; i < selectedSlides.Count; i++)
            {
                int row = i / layoutInfo.Columns;
                int col = i % layoutInfo.Columns;

                float x = margin + col * (thumbWidth + margin);
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
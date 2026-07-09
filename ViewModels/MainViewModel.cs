using Microsoft.Win32;
using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;
using TocBuilder_dotnet_framework.Commands;
using TocBuilder_dotnet_framework.Models;
using TocBuilder_dotnet_framework.Services;

namespace TocBuilder_dotnet_framework.ViewModels
{
    public class MainViewModel : INotifyPropertyChanged
    {
        private readonly PreviewThumbnailService _previewthumbnailService;
        private readonly TocGeneratorService _tocService;

        private string _filePath;
        private int _columns = 3;
        private int _margin = 20;
        private string _status;
        private bool _isBusy;
        private bool _useFixedColumns = false;

        private double _previewCanvasWidth = LayoutConstants.DefaultSlideWidth;
        private double _previewCanvasHeight = LayoutConstants.DefaultSlideHeight;

        private float _actualSlideWidth = LayoutConstants.DefaultSlideWidth;
        private float _actualSlideHeight = LayoutConstants.DefaultSlideHeight;

        private double _previewViewportWidth;
        private double _previewViewportHeight;

        private double _previewScale = 1.0;
        private double _autoScale = 1.0;

        private string _selectedFont;
        public string SelectedFont
        {
            get => _selectedFont;
            set { _selectedFont = value; OnPropertyChanged(); }
        }

        public ObservableCollection<string> AvailableFonts { get; } = new ObservableCollection<string>();

        private int _selectedFontSize = 12;
        public int SelectedFontSize
        {
            get => _selectedFontSize;
            set { _selectedFontSize = value; OnPropertyChanged(); }
        }

        public ObservableCollection<int> AvailableFontSizes { get; } = new ObservableCollection<int>();

        public double PreviewCanvasWidth
        {
            get => _previewCanvasWidth;
            set { _previewCanvasWidth = value; OnPropertyChanged(); }
        }

        public double PreviewCanvasHeight
        {
            get => _previewCanvasHeight;
            set { _previewCanvasHeight = value; OnPropertyChanged(); }
        }

        public byte[] PreviewBackGround
        {
            get
            {
                var bg = Slides.FirstOrDefault(s => s.IsBackground);
                return bg?.Thumbnail;
            }
        }

        public double PreviewScale
        {
            get => _previewScale;
            set { _previewScale = value; OnPropertyChanged(); }
        }

        public double AutoScale
        {
            get => _autoScale;
            set { _autoScale = value; OnPropertyChanged(); }
        }

        public float ActualSlideWidth => _actualSlideWidth;
        public float ActualSlideHeight => _actualSlideHeight;

        public string FilePath
        {
            get => _filePath;
            set
            {
                if (_filePath != value)
                {
                    _filePath = value;
                    OnPropertyChanged();
                    _ = LoadSlidesAsync();
                    UpdateCanGenerate();
                }
            }
        }

        public int Columns
        {
            get => _columns;
            set
            {
                if (_columns != value)
                {
                    _columns = value;
                    OnPropertyChanged();
                    if (UseFixedColumns) UpdatePreview();
                    OnPropertyChanged(nameof(ColumnsDisplayText));
                }
            }
        }

        public string ColumnsDisplayText
        {
            get => UseFixedColumns ? Columns.ToString() : "авто";
        }

        public int Margin
        {
            get => _margin;
            set { _margin = value; OnPropertyChanged(); UpdatePreview(); }
        }

        public bool UseFixedColumns
        {
            get => _useFixedColumns;
            set
            {
                if (_useFixedColumns != value)
                {
                    _useFixedColumns = value;
                    OnPropertyChanged();
                    OnPropertyChanged(nameof(ColumnsDisplayText));
                    UpdatePreview();
                }
            }
        }

        public string Status
        {
            get => _status;
            set { _status = value; OnPropertyChanged(); }
        }

        public bool IsBusy
        {
            get => _isBusy;
            set { _isBusy = value; OnPropertyChanged(); UpdateCanGenerate(); }
        }

        public int SelectedBackgroundSlideIndex =>
            Slides
                .Select((s, i) => new { Slide = s, Index = i })
                .FirstOrDefault(x => x.Slide.IsBackground)
                ?.Index
            ?? 0; // fallback: первый слайд

        public bool CanGenerate => !IsBusy && Slides.Any(s => s.IsSelected);

        public ObservableCollection<SlideItem> Slides { get; } = new ObservableCollection<SlideItem>();
        public ObservableCollection<PreviewItem> PreviewItems { get; } = new ObservableCollection<PreviewItem>();

        public ICommand BrowseCommand { get; }
        public ICommand GenerateCommand { get; }
        public ICommand SelectAllCommand { get; }
        public ICommand DeselectAllCommand { get; }

        public MainViewModel()
        {
            _previewthumbnailService = new PreviewThumbnailService();
            _tocService = new TocGeneratorService();

            Status = "Выберите презентацию";

            // Загрузка системных шрифтов
            var systemFonts = System.Windows.Media.Fonts.SystemFontFamilies
                .Select(f => f.Source)
                .OrderBy(name => name)
                .ToList();

            foreach (var font in systemFonts)
            {
                AvailableFonts.Add(font);
            }

            // Шрифт по умолчанию
            SelectedFont = AvailableFonts.Contains("Calibri") ? "Calibri" : (AvailableFonts.FirstOrDefault() ?? "Arial");

            // Загрузка размеров шрифтов
            var fontSizes = new[] { 8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 32, 36, 40 };
            foreach (var size in fontSizes)
            {
                AvailableFontSizes.Add(size);
            }
            SelectedFontSize = 12;

            BrowseCommand = new RelayCommand(BrowseFile);
            GenerateCommand = new AsyncRelayCommand(async () => await GenerateTocAsync(), () => CanGenerate);
            SelectAllCommand = new RelayCommand(() => SelectAll(true));
            DeselectAllCommand = new RelayCommand(() => SelectAll(false));

            PropertyChanged += MainViewModel_PropertyChanged;
        }

        private void MainViewModel_PropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            if (e.PropertyName == nameof(Margin) || e.PropertyName == nameof(Columns))
            {
                UpdatePreview();
            }
        }

        private void Slide_PropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            if (e.PropertyName == nameof(SlideItem.IsSelected))
            {
                UpdatePreview();
                UpdateCanGenerate();
            }
            else if (e.PropertyName == nameof(SlideItem.IsBackground))
            {
                var changedSlide = sender as SlideItem;
                if (changedSlide.IsBackground)
                {
                    // Снимаем IsBackground с других слайдов
                    foreach (var slide in Slides)
                    {
                        if (slide != changedSlide && slide.IsBackground)
                            slide.IsBackground = false;
                    }
                }

                // Обновляем фон превью
                OnPropertyChanged(nameof(PreviewBackGround));
                UpdatePreview();
            }
        }

        private void BrowseFile()
        {
            var dlg = new OpenFileDialog { Filter = "PowerPoint презентации|*.pptx;*.ppt", Title = "Выберите презентацию" };
            if (dlg.ShowDialog() == true)
            {
                FilePath = dlg.FileName;
            }
        }

        private async Task LoadSlidesAsync()
        {
            Slides.Clear();
            if (string.IsNullOrEmpty(FilePath) || !File.Exists(FilePath)) return;

            IsBusy = true;
            Status = "Загрузка слайдов...";

            try
            {
                var slides = await Task.Run(() => _previewthumbnailService.GetSlides(FilePath));
                var dimensions = await Task.Run(() => _previewthumbnailService.GetSlideDimensions(FilePath));

                (_actualSlideWidth, _actualSlideHeight) = dimensions;

                foreach (var slide in slides)
                {
                    slide.PropertyChanged += Slide_PropertyChanged;
                    Slides.Add(slide);
                }

                var themeFonts = await Task.Run(() => _previewthumbnailService.GetThemeFonts(FilePath));
                if (themeFonts.Any())
                {
                    string themeBodyFont = themeFonts.First();
                    if (AvailableFonts.Contains(themeBodyFont))
                    {
                        SelectedFont = themeBodyFont;
                    }
                }

                Status = $"Загружено {Slides.Count} слайдов";
                UpdatePreview();
                UpdateCanGenerate();
            }
            catch (Exception ex)
            {
                Status = $"Ошибка: {ex.Message}";
                MessageBox.Show($"Не удалось загрузить презентацию:\n{ex.Message}", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                IsBusy = false;
            }
        }

        private void UpdatePreview()
        {
            IsBusy = true;
            Status = "Обновление превью...";

            var selectedSlides = Slides.Where(s => s.IsSelected).ToList();
            if (!selectedSlides.Any())
            {
                PreviewItems.Clear();
                IsBusy = false;
                return;
            }

            int columnsForCalc = UseFixedColumns ? Columns : -1;

            var previewItems = LayoutCalculatorService.GeneratePreviewItems(
                selectedSlides,
                columnsForCalc,
                Margin,
                ActualSlideWidth,
                ActualSlideHeight);

            Application.Current.Dispatcher.Invoke(() =>
            {
                PreviewItems.Clear();

                PreviewCanvasWidth = ActualSlideWidth;
                PreviewCanvasHeight = ActualSlideHeight;

                foreach (var item in previewItems)
                    PreviewItems.Add(item);
            });

            CalculateAutoScale();

            IsBusy = false;
            Status = "Превью обновлено";
        }

        private void CalculateAutoScale()
        {
            if (_previewViewportWidth <= 0 ||
                _previewViewportHeight <= 0 ||
                ActualSlideWidth <= 0 ||
                ActualSlideHeight <= 0)
                return;

            double scaleX = _previewViewportWidth / ActualSlideWidth;
            double scaleY = _previewViewportHeight / ActualSlideHeight;

            AutoScale = Math.Min(scaleX, scaleY);
            AutoScale = Math.Max(0.3, Math.Min(1.2, AutoScale));
            PreviewScale = AutoScale;
        }

        public void UpdatePreviewViewportSize(double width, double height)
        {
            _previewViewportWidth = width;
            _previewViewportHeight = height;

            CalculateAutoScale();
        }

        private async Task GenerateTocAsync()
        {
            if (!Slides.Any(s => s.IsSelected)) return;

            IsBusy = true;
            Status = "Создание оглавления...";

            try
            {
                var selectedSlides = Slides.Where(s => s.IsSelected).ToList();

                string outputPath = await Task.Run(() => _tocService.CreateTableOfContents(FilePath, selectedSlides, Columns, Margin, SelectedBackgroundSlideIndex, SelectedFont, SelectedFontSize));

                Status = $"✅ Готово! Файл сохранён: {Path.GetFileName(outputPath)}";

                if (MessageBox.Show($"Оглавление создано!\n\nФайл: {outputPath}\n\nОткрыть презентацию?", "Готово", MessageBoxButton.YesNo, MessageBoxImage.Information) == MessageBoxResult.Yes)
                    System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(outputPath) { UseShellExecute = true });
            }
            catch (Exception ex)
            {
                Status = $"❌ Ошибка: {ex.Message}";
                MessageBox.Show($"Не удалось создать оглавление:\n{ex.Message}", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                IsBusy = false;
                UpdateCanGenerate();
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        private void SelectAll(bool select)
        {
            foreach (var slide in Slides) slide.IsSelected = select;
            UpdateCanGenerate();
            UpdatePreview(); // Один вызов после всех изменений
        }

        private void UpdateCanGenerate()
        {
            OnPropertyChanged(nameof(CanGenerate));
            if (GenerateCommand is AsyncRelayCommand cmd) cmd.RaiseCanExecuteChanged();
        }

        #region INotifyPropertyChanged
        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged([CallerMemberName] string propName = null) =>
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propName));
        #endregion
    }
}
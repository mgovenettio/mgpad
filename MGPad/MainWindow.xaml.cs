using System;
using System.ComponentModel;
using System.IO;
using System.Globalization;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Input;
using System.Text;
using System.Linq;
using Microsoft.Win32;
using System.Windows.Threading;
using System.Windows.Media;
using PdfSharp.Drawing;
using PdfSharp.Pdf;

namespace MGPad;

public enum DocumentType
{
    RichText,
    PlainText,
    Markdown
}

public partial class MainWindow : Window
{
    public static readonly RoutedUICommand ToggleBoldCommand =
        new RoutedUICommand("ToggleBold", "ToggleBold", typeof(MainWindow),
            new InputGestureCollection { new KeyGesture(Key.B, ModifierKeys.Control) });

    public static readonly RoutedUICommand ToggleUnderlineCommand =
        new RoutedUICommand("ToggleUnderline", "ToggleUnderline", typeof(MainWindow),
            new InputGestureCollection { new KeyGesture(Key.U, ModifierKeys.Control) });

    public static readonly RoutedUICommand ToggleInputLanguageCommand =
        new RoutedUICommand("ToggleInputLanguage", "ToggleInputLanguage", typeof(MainWindow),
            new InputGestureCollection { new KeyGesture(Key.J, ModifierKeys.Control) });

    public static readonly RoutedUICommand InsertTimestampCommand =
        new RoutedUICommand("InsertTimestamp", "InsertTimestamp", typeof(MainWindow),
            new InputGestureCollection { new KeyGesture(Key.T, ModifierKeys.Control) });

    private string? _currentFilePath;
    private DocumentType _currentDocumentType;
    private bool _isDirty;
    private bool _isLoadingDocument;
    private bool _allowCloseWithoutPrompt;
    private bool _isMarkdownMode = false;
    private bool _isNightMode = false;
    private const double ZoomStep = 0.1;
    private const double MinZoom = 0.5;
    private const double MaxZoom = 3.0;
    private double _zoomLevel = 1.0;
    private const double MinFontSize = 6;
    private const double MaxFontSize = 96;
    private const int MaxRecentDocuments = 10;
    private readonly List<string> _recentDocuments = new();
    private readonly string _recentDocumentsFilePath;
    private readonly DispatcherTimer _markdownPreviewTimer;
    private CultureInfo? _englishInputLanguage;
    private CultureInfo? _japaneseInputLanguage;
    private bool _isUpdatingFontControls;
    private readonly double[] _defaultFontSizes = new double[]
        { 8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72 };

    private sealed class PdfTextRun
    {
        public string Text { get; set; } = string.Empty;
        public bool IsBold { get; set; }
        public bool IsUnderline { get; set; }
    }

    private sealed class PdfParagraph
    {
        public List<PdfTextRun> Runs { get; } = new();
    }

    private sealed class PdfLineSpan
    {
        public string Text { get; set; } = string.Empty;
        public XFont Font { get; set; } = null!;
    }

    public MainWindow()
    {
        InitializeComponent();
        _recentDocumentsFilePath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "MGPad",
            "recent-documents.txt");
        _markdownPreviewTimer = new DispatcherTimer
        {
            Interval = TimeSpan.FromMilliseconds(300)
        };
        _markdownPreviewTimer.Tick += (s, e) =>
        {
            _markdownPreviewTimer.Stop();
            UpdateMarkdownPreview();
        };
        InputLanguageManager.Current.InputLanguageChanged += InputLanguageManager_InputLanguageChanged;
        CommandBindings.Add(new CommandBinding(ToggleBoldCommand,
            (s, e) =>
            {
                if (!CanFormat())
                {
                    e.Handled = true;
                    return;
                }

                ToggleBold();
                MarkDirty();
                e.Handled = true;
            }));

        CommandBindings.Add(new CommandBinding(ToggleUnderlineCommand,
            (s, e) =>
            {
                if (!CanFormat())
                {
                    e.Handled = true;
                    return;
                }

                ToggleUnderline();
                MarkDirty();
                e.Handled = true;
            }));
        CommandBindings.Add(new CommandBinding(
            ToggleInputLanguageCommand,
            (s, e) => ToggleInputLanguage()));
        CommandBindings.Add(new CommandBinding(
            InsertTimestampCommand,
            (s, e) => InsertTimestampAtCaret()));
        InitializePreferredInputLanguages();

        PopulateFontControls();

        if (EditorBox != null)
        {
            EditorBox.SelectionChanged += EditorBox_SelectionChanged;
        }

        if (_englishInputLanguage == null || _japaneseInputLanguage == null)
        {
            if (ToggleLanguageButton != null)
                ToggleLanguageButton.IsEnabled = false;
        }
        else
        {
            if (ToggleLanguageButton != null)
                ToggleLanguageButton.IsEnabled = true;
        }

        InitializeDocumentState();
        UpdateFormattingControls();
        UpdateLanguageIndicator();

        SetMarkdownMode(false);
        ApplyTheme();
        ApplyZoom();
        LoadRecentDocuments();
        UpdateRecentDocumentsMenu();
    }

    private void ApplyMarkdownModeLayout()
    {
        if (EditorRegionGrid == null)
            return;

        if (_isMarkdownMode)
        {
            EditorRegionGrid.ColumnDefinitions[0].Width = new GridLength(3, GridUnitType.Star);
            EditorRegionGrid.ColumnDefinitions[1].Width = new GridLength(2, GridUnitType.Star);
            if (MarkdownPreviewContainer != null)
                MarkdownPreviewContainer.Visibility = Visibility.Visible;
        }
        else
        {
            EditorRegionGrid.ColumnDefinitions[0].Width = new GridLength(1, GridUnitType.Star);
            EditorRegionGrid.ColumnDefinitions[1].Width = new GridLength(0);
            if (MarkdownPreviewContainer != null)
                MarkdownPreviewContainer.Visibility = Visibility.Collapsed;
        }
    }

    private void SetMarkdownMode(bool enable)
    {
        _isMarkdownMode = enable;

        if (MarkdownModeMenuItem != null)
            MarkdownModeMenuItem.IsChecked = enable;

        ApplyMarkdownModeLayout();
        UpdateMarkdownPreview();
    }

    private void MarkdownModeMenuItem_Click(object sender, RoutedEventArgs e)
    {
        bool enable = MarkdownModeMenuItem?.IsChecked ?? false;
        SetMarkdownMode(enable);
    }

    private void NightModeButton_Click(object sender, RoutedEventArgs e)
    {
        _isNightMode = !_isNightMode;
        ApplyTheme();
    }

    private void ApplyTheme()
    {
        Brush windowBackground = _isNightMode
            ? new SolidColorBrush(Color.FromRgb(30, 30, 30))
            : Brushes.White;
        Brush panelBackground = _isNightMode
            ? new SolidColorBrush(Color.FromRgb(45, 45, 48))
            : Brushes.WhiteSmoke;
        Brush borderBrush = _isNightMode
            ? new SolidColorBrush(Color.FromRgb(62, 62, 66))
            : Brushes.LightGray;
        Brush foreground = _isNightMode ? Brushes.WhiteSmoke : Brushes.Black;

        if (RootGrid != null)
            RootGrid.Background = windowBackground;

        if (EditorBorder != null)
            EditorBorder.BorderBrush = borderBrush;

        if (EditorBox != null)
        {
            EditorBox.Background = panelBackground;
            EditorBox.Foreground = foreground;
            EditorBox.BorderBrush = borderBrush;
            if (EditorBox.Document != null)
                EditorBox.Document.Foreground = foreground;
        }

        if (MarkdownPreviewContainer != null)
        {
            MarkdownPreviewContainer.Background = panelBackground;
            MarkdownPreviewContainer.BorderBrush = borderBrush;
        }

        if (MarkdownPreviewTextBlock != null)
            MarkdownPreviewTextBlock.Foreground = foreground;

        if (MainStatusBar != null)
        {
            MainStatusBar.Background = panelBackground;
            MainStatusBar.Foreground = foreground;
        }

        if (ZoomPercentageTextBlock != null)
            ZoomPercentageTextBlock.Foreground = foreground;

        if (NightModeButton != null)
            NightModeButton.Content = _isNightMode ? "Day" : "Night";
    }

    private void ApplyZoom()
    {
        if (EditorBox != null)
            EditorBox.LayoutTransform = new ScaleTransform(_zoomLevel, _zoomLevel);

        if (ZoomPercentageTextBlock != null)
            ZoomPercentageTextBlock.Text = $"{Math.Round(_zoomLevel * 100)}%";

        if (ZoomInButton != null)
            ZoomInButton.IsEnabled = _zoomLevel < MaxZoom;

        if (ZoomOutButton != null)
            ZoomOutButton.IsEnabled = _zoomLevel > MinZoom;
    }

    private void ChangeZoom(double delta)
    {
        double newLevel = _zoomLevel + delta;
        newLevel = Math.Max(MinZoom, Math.Min(MaxZoom, newLevel));

        if (Math.Abs(newLevel - _zoomLevel) < 0.0001)
            return;

        _zoomLevel = newLevel;
        ApplyZoom();
    }

    private void ZoomInButton_Click(object sender, RoutedEventArgs e)
    {
        ChangeZoom(ZoomStep);
    }

    private void ZoomOutButton_Click(object sender, RoutedEventArgs e)
    {
        ChangeZoom(-ZoomStep);
    }

    private void UpdateMarkdownPreview()
    {
        if (!_isMarkdownMode || MarkdownPreviewTextBlock == null)
            return;

        string markdown = GetEditorPlainText();
        if (string.IsNullOrWhiteSpace(markdown))
        {
            MarkdownPreviewTextBlock.Text = string.Empty;
            return;
        }

        var lines = markdown.Replace("\r\n", "\n").Split('\n');
        var sb = new StringBuilder();

        bool inCodeBlock = false;

        foreach (var rawLine in lines)
        {
            string line = rawLine;

            // Detect fenced code block markers
            if (line.TrimStart().StartsWith("```"))
            {
                inCodeBlock = !inCodeBlock;
                sb.AppendLine(line);
                continue;
            }

            if (inCodeBlock)
            {
                sb.AppendLine("    " + line);
                continue;
            }

            // Simple headings
            if (line.StartsWith("# "))
            {
                string text = line.Substring(2).Trim();
                sb.AppendLine(text.ToUpperInvariant());
                sb.AppendLine(new string('=', Math.Max(text.Length, 3)));
                sb.AppendLine();
                continue;
            }
            if (line.StartsWith("## "))
            {
                string text = line.Substring(3).Trim();
                sb.AppendLine(text);
                sb.AppendLine(new string('-', Math.Max(text.Length, 3)));
                sb.AppendLine();
                continue;
            }
            if (line.StartsWith("### "))
            {
                string text = line.Substring(4).Trim();
                sb.AppendLine("### " + text);
                continue;
            }

            // Bullet lists (keep as-is)
            if (line.TrimStart().StartsWith("- ") || line.TrimStart().StartsWith("* "))
            {
                sb.AppendLine(line);
                continue;
            }

            // Simple inline emphasis: leave **bold** and *italic* as-is for now
            sb.AppendLine(line);
        }

        MarkdownPreviewTextBlock.Text = sb.ToString();
    }

    private void SetEditorPlainText(string text)
    {
        if (EditorBox == null)
            return;

        EditorBox.Document = new FlowDocument();
        var range = new TextRange(EditorBox.Document.ContentStart, EditorBox.Document.ContentEnd);
        range.Text = text;
        ApplyTheme();
    }

    private void LoadRtfIntoEditor(string path)
    {
        if (EditorBox == null)
            return;

        EditorBox.Document = new FlowDocument();
        var range = new TextRange(EditorBox.Document.ContentStart, EditorBox.Document.ContentEnd);

        using var stream = new FileStream(path, FileMode.Open, FileAccess.Read);
        range.Load(stream, DataFormats.Rtf);
        ApplyTheme();
    }

    private string GetEditorPlainText()
    {
        if (EditorBox == null)
            return string.Empty;

        TextRange range = new TextRange(EditorBox.Document.ContentStart, EditorBox.Document.ContentEnd);
        return range.Text ?? string.Empty;
    }

    private void ScheduleMarkdownPreviewUpdate()
    {
        if (!_isMarkdownMode)
            return;

        _markdownPreviewTimer.Stop();
        _markdownPreviewTimer.Start();
    }

    private void InsertTimestampAtCaret()
    {
        if (EditorBox == null)
            return;

        string timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm");

        TextSelection selection = EditorBox.Selection;

        if (!selection.IsEmpty)
        {
            selection.Text = timestamp;
            EditorBox.CaretPosition = selection.End;
        }
        else
        {
            var caret = EditorBox.CaretPosition;

            if (!caret.IsAtInsertionPosition)
                caret = caret.GetInsertionPosition(LogicalDirection.Forward);

            selection = new TextRange(caret, caret);
            selection.Text = timestamp;

            EditorBox.CaretPosition = selection.End;
        }

        EditorBox.Focus();
        MarkDirty();
        ScheduleMarkdownPreviewUpdate();
    }

    private void InitializePreferredInputLanguages()
    {
        _englishInputLanguage = null;
        _japaneseInputLanguage = null;

        var available = InputLanguageManager.Current.AvailableInputLanguages;
        if (available == null)
            return;

        foreach (var langObj in available)
        {
            if (langObj is CultureInfo ci)
            {
                string twoLetter = ci.TwoLetterISOLanguageName.ToLowerInvariant();
                if (twoLetter == "en" && _englishInputLanguage == null)
                {
                    _englishInputLanguage = ci;
                }
                else if (twoLetter == "ja" && _japaneseInputLanguage == null)
                {
                    _japaneseInputLanguage = ci;
                }
            }
        }
    }

    private void InitializeDocumentState()
    {
        _isDirty = false;
        SetCurrentFile(null, DocumentType.RichText);
        MarkClean();
    }

    private void SetCurrentFile(string? path, DocumentType type)
    {
        _currentFilePath = path;
        _currentDocumentType = type;
        if (!string.IsNullOrWhiteSpace(path))
        {
            AddRecentDocument(path);
        }
        UpdateWindowTitle();
    }

    private void MarkDirty()
    {
        if (_isDirty)
        {
            return;
        }

        _isDirty = true;
        UpdateWindowTitle();
    }

    private void MarkClean()
    {
        if (!_isDirty)
        {
            UpdateWindowTitle();
            return;
        }

        _isDirty = false;
        UpdateWindowTitle();
    }

    private void UpdateWindowTitle()
    {
        var fileName = GetDisplayFileName();
        var dirtyMarker = _isDirty ? "*" : string.Empty;
        Title = $"MGPad - {fileName}{dirtyMarker}";
        UpdateStatusBar();
    }

    private string GetDisplayFileName()
    {
        return string.IsNullOrEmpty(_currentFilePath)
            ? "Untitled"
            : Path.GetFileName(_currentFilePath);
    }

    private void UpdateStatusBar()
    {
        if (FileNameTextBlock is null)
        {
            return;
        }

        var displayName = GetDisplayFileName();
        if (_isDirty)
        {
            displayName += " *";
        }

        FileNameTextBlock.Text = displayName;
        UpdateLanguageIndicator();
    }

    private void LoadRecentDocuments()
    {
        _recentDocuments.Clear();

        try
        {
            if (!File.Exists(_recentDocumentsFilePath))
            {
                return;
            }

            foreach (var line in File.ReadAllLines(_recentDocumentsFilePath))
            {
                var path = line.Trim();
                if (string.IsNullOrWhiteSpace(path))
                    continue;

                if (File.Exists(path) && !_recentDocuments.Any(p => string.Equals(p, path, StringComparison.OrdinalIgnoreCase)))
                {
                    _recentDocuments.Add(path);
                    if (_recentDocuments.Count >= MaxRecentDocuments)
                        break;
                }
            }
        }
        catch
        {
            // Ignore errors loading recent documents to avoid disrupting startup.
        }
    }

    private void SaveRecentDocuments()
    {
        try
        {
            var directory = Path.GetDirectoryName(_recentDocumentsFilePath);
            if (!string.IsNullOrWhiteSpace(directory))
            {
                Directory.CreateDirectory(directory);
            }

            File.WriteAllLines(_recentDocumentsFilePath, _recentDocuments);
        }
        catch
        {
            // Ignore errors saving recent documents to avoid interrupting the user.
        }
    }

    private void UpdateRecentDocumentsMenu()
    {
        if (RecentDocumentsMenuItem == null)
            return;

        RecentDocumentsMenuItem.Items.Clear();

        if (_recentDocuments.Count == 0)
        {
            RecentDocumentsMenuItem.Items.Add(new MenuItem
            {
                Header = "No recent documents",
                IsEnabled = false
            });
            return;
        }

        foreach (var path in _recentDocuments)
        {
            var item = new MenuItem
            {
                Header = $"{Path.GetFileName(path)} ({path})",
                Tag = path
            };

            item.Click += RecentDocumentMenuItem_Click;
            RecentDocumentsMenuItem.Items.Add(item);
        }
    }

    private void AddRecentDocument(string path)
    {
        if (string.IsNullOrWhiteSpace(path))
            return;

        string fullPath = Path.GetFullPath(path);
        _recentDocuments.RemoveAll(p => string.Equals(p, fullPath, StringComparison.OrdinalIgnoreCase));
        _recentDocuments.Insert(0, fullPath);

        if (_recentDocuments.Count > MaxRecentDocuments)
        {
            _recentDocuments.RemoveRange(MaxRecentDocuments, _recentDocuments.Count - MaxRecentDocuments);
        }

        SaveRecentDocuments();
        UpdateRecentDocumentsMenu();
    }

    private void RemoveRecentDocument(string path)
    {
        bool removed = _recentDocuments.RemoveAll(p => string.Equals(p, path, StringComparison.OrdinalIgnoreCase)) > 0;
        if (removed)
        {
            SaveRecentDocuments();
            UpdateRecentDocumentsMenu();
        }
    }

    private void UpdateLanguageIndicator()
    {
        var current = InputLanguageManager.Current.CurrentInputLanguage;
        if (LanguageIndicatorTextBlock == null || current == null)
        {
            return;
        }

        string twoLetter = current.TwoLetterISOLanguageName.ToLowerInvariant();
        string code;

        switch (twoLetter)
        {
            case "en":
                code = "EN";
                break;
            case "ja":
                code = "JP";
                break;
            default:
                code = twoLetter.ToUpperInvariant();
                break;
        }

        LanguageIndicatorTextBlock.Text = code;
    }

    private void InputLanguageManager_InputLanguageChanged(object? sender, InputLanguageEventArgs e)
    {
        UpdateLanguageIndicator();
    }

    private void ToggleLanguageButton_Click(object sender, RoutedEventArgs e)
    {
        ToggleInputLanguage();
    }

    private void InsertTimestampMenuItem_Click(object sender, RoutedEventArgs e)
    {
        InsertTimestampAtCaret();
    }

    private void TimestampButton_Click(object sender, RoutedEventArgs e)
    {
        InsertTimestampAtCaret();
    }

    private void ToggleInputLanguage()
    {
        if (_englishInputLanguage == null || _japaneseInputLanguage == null)
        {
            UpdateLanguageIndicator();
            return;
        }

        var current = InputLanguageManager.Current.CurrentInputLanguage;

        if (current == null)
        {
            UpdateLanguageIndicator();
            return;
        }

        // Decide which language to switch to
        CultureInfo target;
        string currentTwo = current.TwoLetterISOLanguageName.ToLowerInvariant();

        if (currentTwo == "ja")
        {
            target = _englishInputLanguage;
        }
        else if (currentTwo == "en")
        {
            target = _japaneseInputLanguage;
        }
        else
        {
            // If current isn't EN or JA, default to EN
            target = _englishInputLanguage;
        }

        InputLanguageManager.Current.CurrentInputLanguage = target;
        UpdateLanguageIndicator();
    }

    private void FileNew_Click(object sender, RoutedEventArgs e) => CreateNewDocument();

    private void FileOpen_Click(object sender, RoutedEventArgs e) => OpenDocumentFromDialog();

    private void FileSave_Click(object sender, RoutedEventArgs e) => SaveCurrentDocument();

    private void FileSaveAs_Click(object sender, RoutedEventArgs e) => SaveDocumentWithDialog();

    private void ExportPdfMenuItem_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            var dialog = new Microsoft.Win32.SaveFileDialog
            {
                Title = "Export as PDF",
                Filter = "PDF Files (*.pdf)|*.pdf|All Files (*.*)|*.*",
                DefaultExt = ".pdf",
                FileName = !string.IsNullOrEmpty(_currentFilePath)
                    ? Path.GetFileNameWithoutExtension(_currentFilePath) + ".pdf"
                    : "Document.pdf"
            };

            bool? result = dialog.ShowDialog(this);
            if (result != true)
                return;

            string pdfPath = dialog.FileName;
            ExportDocumentToPdf(pdfPath);

            MessageBox.Show(this,
                "PDF export completed:\n" + pdfPath,
                "Export as PDF",
                MessageBoxButton.OK,
                MessageBoxImage.Information);
        }
        catch (Exception ex)
        {
            MessageBox.Show(this,
                "Failed to export PDF.\n\n" + ex.Message,
                "Export as PDF",
                MessageBoxButton.OK,
                MessageBoxImage.Error);
        }
    }

    private List<PdfParagraph> ExtractParagraphsForPdf()
    {
        var result = new List<PdfParagraph>();

        if (EditorBox?.Document == null)
            return result;

        foreach (var block in EditorBox.Document.Blocks)
        {
            CollectParagraphs(block, result);
        }

        return result;
    }

    private void CollectParagraphs(Block block, List<PdfParagraph> result, string? prefix = null)
    {
        switch (block)
        {
            case Paragraph paragraph:
                var pdfParagraph = CreatePdfParagraphFromParagraph(paragraph, prefix);
                if (pdfParagraph != null)
                    result.Add(pdfParagraph);
                break;
            case List list:
                int index = list.StartIndex <= 0 ? 1 : list.StartIndex;
                foreach (ListItem item in list.ListItems)
                {
                    string marker = GetListMarkerText(list.MarkerStyle, index);
                    CollectParagraphsFromListItem(item, result, marker);
                    index++;
                }
                break;
            case Section section:
                foreach (var child in section.Blocks)
                {
                    CollectParagraphs(child, result, prefix);
                }
                break;
        }
    }

    private void CollectParagraphsFromListItem(ListItem item, List<PdfParagraph> result, string marker)
    {
        bool isFirstBlock = true;
        foreach (var child in item.Blocks)
        {
            string? prefix = isFirstBlock ? marker : new string(' ', marker.Length);
            CollectParagraphs(child, result, prefix);
            isFirstBlock = false;
        }
    }

    private PdfParagraph? CreatePdfParagraphFromParagraph(Paragraph paragraph, string? prefix)
    {
        var pdfParagraph = new PdfParagraph();

        if (!string.IsNullOrEmpty(prefix))
        {
            pdfParagraph.Runs.Add(new PdfTextRun { Text = prefix });
        }

        foreach (Inline inline in paragraph.Inlines)
        {
            if (inline is Run run)
            {
                string text = run.Text;
                if (string.IsNullOrEmpty(text))
                    continue;

                bool isBold = run.FontWeight == FontWeights.Bold;
                bool isUnderline = run.TextDecorations?.Contains(TextDecorations.Underline[0]) == true;

                pdfParagraph.Runs.Add(new PdfTextRun
                {
                    Text = text,
                    IsBold = isBold,
                    IsUnderline = isUnderline
                });
            }
            else if (inline is LineBreak)
            {
                // Represent line breaks as separate runs if desired
                pdfParagraph.Runs.Add(new PdfTextRun { Text = "\n" });
            }
        }

        // Ignore completely empty paragraphs
        return pdfParagraph.Runs.Count > 0 ? pdfParagraph : null;
    }

    private string GetListMarkerText(TextMarkerStyle style, int index)
    {
        return style switch
        {
            TextMarkerStyle.None => string.Empty,
            TextMarkerStyle.Disc => "• ",
            TextMarkerStyle.Circle => "○ ",
            TextMarkerStyle.Square => "▪ ",
            TextMarkerStyle.Box => "▪ ",
            TextMarkerStyle.LowerRoman => $"{ToRoman(index).ToLowerInvariant()}. ",
            TextMarkerStyle.UpperRoman => $"{ToRoman(index)}. ",
            TextMarkerStyle.LowerLatin => $"{ToAlphabetic(index, false)}. ",
            TextMarkerStyle.UpperLatin => $"{ToAlphabetic(index, true)}. ",
            _ => $"{index}. ",
        };
    }

    private string ToAlphabetic(int number, bool upper)
    {
        if (number <= 0)
            return string.Empty;

        var builder = new StringBuilder();
        int n = number;
        while (n > 0)
        {
            n--;
            builder.Insert(0, (char)((n % 26) + (upper ? 'A' : 'a')));
            n /= 26;
        }

        return builder.ToString();
    }

    private string ToRoman(int number)
    {
        if (number <= 0)
            return string.Empty;

        (int value, string symbol)[] map =
        {
            (1000, "M"), (900, "CM"), (500, "D"), (400, "CD"),
            (100, "C"), (90, "XC"), (50, "L"), (40, "XL"),
            (10, "X"), (9, "IX"), (5, "V"), (4, "IV"), (1, "I")
        };

        var result = new StringBuilder();
        int remaining = number;

        foreach (var (value, symbol) in map)
        {
            while (remaining >= value)
            {
                result.Append(symbol);
                remaining -= value;
            }
        }

        return result.ToString();
    }

    private void ExportDocumentToPdf(string pdfPath)
    {
        var paragraphs = ExtractParagraphsForPdf();

        var document = new PdfDocument();
        document.Info.Title = string.IsNullOrEmpty(_currentFilePath)
            ? "MGPad Document"
            : Path.GetFileNameWithoutExtension(_currentFilePath);

        PdfPage page = document.AddPage();
        XGraphics gfx = XGraphics.FromPdfPage(page);

        // Basic fonts
        XFont regularFont = new XFont("Segoe UI", 12, XFontStyle.Regular);
        XFont boldFont = new XFont("Segoe UI", 12, XFontStyle.Bold);
        XFont underlineFont = new XFont("Segoe UI", 12, XFontStyle.Underline);
        XFont boldUnderlineFont = new XFont("Segoe UI", 12, XFontStyle.Bold | XFontStyle.Underline);

        // Layout: 1-inch margins (72 points per inch)
        double marginLeft = 72;
        double marginTop = 72;
        double marginRight = 72;
        double marginBottom = 72;

        double y = marginTop;
        foreach (var paragraph in paragraphs)
        {
            // For simplicity, choose a font based on the first run that has formatting
            bool anyBold = paragraph.Runs.Any(r => r.IsBold);
            bool anyUnderline = paragraph.Runs.Any(r => r.IsUnderline);

            XFont fontToUse = regularFont;
            if (anyBold && anyUnderline)
                fontToUse = boldUnderlineFont;
            else if (anyBold)
                fontToUse = boldFont;
            else if (anyUnderline)
                fontToUse = underlineFont;

            double lineHeight = fontToUse.GetHeight(gfx) * 1.4;
            double usableWidth = page.Width - marginLeft - marginRight;

            var wrappedLines = WrapParagraphRuns(
                paragraph.Runs,
                gfx,
                usableWidth,
                regularFont,
                boldFont,
                underlineFont,
                boldUnderlineFont);

            foreach (var line in wrappedLines)
            {
                if (y + lineHeight > page.Height - marginBottom)
                {
                    // New page
                    page = document.AddPage();
                    gfx = XGraphics.FromPdfPage(page);
                    y = marginTop;
                }

                if (line.Count == 0)
                {
                    y += lineHeight;
                    continue;
                }

                double x = marginLeft;
                foreach (var span in line)
                {
                    double spanWidth = gfx.MeasureString(span.Text, span.Font).Width;
                    gfx.DrawString(span.Text, span.Font, XBrushes.Black,
                        new XRect(x, y, spanWidth, lineHeight),
                        XStringFormats.TopLeft);
                    x += spanWidth;
                }

                y += lineHeight;
            }

            // Extra space between paragraphs
            y += lineHeight * 0.5;
        }

        document.Save(pdfPath);
    }

    private List<List<PdfLineSpan>> WrapParagraphRuns(
        IEnumerable<PdfTextRun> runs,
        XGraphics gfx,
        double maxWidth,
        XFont regularFont,
        XFont boldFont,
        XFont underlineFont,
        XFont boldUnderlineFont)
    {
        var lines = new List<List<PdfLineSpan>>();
        var currentLine = new List<PdfLineSpan>();
        double currentWidth = 0;

        void AddLine()
        {
            lines.Add(currentLine);
            currentLine = new List<PdfLineSpan>();
            currentWidth = 0;
        }

        void AppendSpan(string text, XFont font)
        {
            if (string.IsNullOrEmpty(text))
                return;

            double width = gfx.MeasureString(text, font).Width;

            if (currentLine.Count > 0 && currentLine[^1].Font == font)
            {
                currentLine[^1].Text += text;
            }
            else
            {
                currentLine.Add(new PdfLineSpan { Text = text, Font = font });
            }

            currentWidth += width;
        }

        int GetFittingLength(string text, XFont font, double availableWidth)
        {
            if (availableWidth <= 0)
                return 0;

            int length = 0;
            for (int i = 1; i <= text.Length; i++)
            {
                double width = gfx.MeasureString(text.Substring(0, i), font).Width;
                if (width <= availableWidth)
                {
                    length = i;
                }
                else
                {
                    break;
                }
            }

            return length;
        }

        foreach (var run in runs)
        {
            XFont font = GetFontForRun(run, regularFont, boldFont, underlineFont, boldUnderlineFont);
            string normalized = run.Text.Replace("\r\n", "\n");
            string[] segments = normalized.Split('\n');

            for (int s = 0; s < segments.Length; s++)
            {
                string segment = segments[s];

                foreach (string token in TokenizeSegment(segment))
                {
                    string remaining = token;
                    while (remaining.Length > 0)
                    {
                        double available = maxWidth - currentWidth;
                        int fitLength = GetFittingLength(remaining, font, available);

                        if (fitLength == 0)
                        {
                            AddLine();
                            available = maxWidth - currentWidth;
                            fitLength = GetFittingLength(remaining, font, available);
                        }

                        string piece = remaining.Substring(0, fitLength);
                        AppendSpan(piece, font);
                        remaining = remaining.Substring(fitLength);

                        if (remaining.Length > 0)
                        {
                            AddLine();
                        }
                    }
                }

                if (s < segments.Length - 1)
                {
                    AddLine();
                }
            }
        }

        if (currentLine.Count > 0 || lines.Count == 0)
        {
            lines.Add(currentLine);
        }

        return lines;
    }

    private static XFont GetFontForRun(PdfTextRun run, XFont regularFont, XFont boldFont, XFont underlineFont, XFont boldUnderlineFont)
    {
        if (run.IsBold && run.IsUnderline)
            return boldUnderlineFont;
        if (run.IsBold)
            return boldFont;
        if (run.IsUnderline)
            return underlineFont;

        return regularFont;
    }

    private IEnumerable<string> TokenizeSegment(string text)
    {
        if (string.IsNullOrEmpty(text))
            yield break;

        var builder = new StringBuilder();
        bool? isWhitespace = null;

        foreach (char c in text)
        {
            bool currentIsWhitespace = char.IsWhiteSpace(c);

            if (isWhitespace == null || currentIsWhitespace == isWhitespace)
            {
                builder.Append(c);
            }
            else
            {
                yield return builder.ToString();
                builder.Clear();
                builder.Append(c);
            }

            isWhitespace = currentIsWhitespace;
        }

        if (builder.Length > 0)
            yield return builder.ToString();
    }

    private void FileExit_Click(object sender, RoutedEventArgs e) => ExitApplication();

    private void FileCommand_CanExecute(object sender, CanExecuteRoutedEventArgs e)
    {
        e.CanExecute = true;
        e.Handled = true;
    }

    private void FileNewCommand_Executed(object sender, ExecutedRoutedEventArgs e)
    {
        CreateNewDocument();
        e.Handled = true;
    }

    private void FileOpenCommand_Executed(object sender, ExecutedRoutedEventArgs e)
    {
        OpenDocumentFromDialog();
        e.Handled = true;
    }

    private void FileSaveCommand_Executed(object sender, ExecutedRoutedEventArgs e)
    {
        SaveCurrentDocument();
        e.Handled = true;
    }

    private void FileSaveAsCommand_Executed(object sender, ExecutedRoutedEventArgs e)
    {
        SaveDocumentWithDialog();
        e.Handled = true;
    }

    private void FileCloseCommand_Executed(object sender, ExecutedRoutedEventArgs e)
    {
        ExitApplication();
        e.Handled = true;
    }

    private void CreateNewDocument()
    {
        if (!ConfirmDiscardUnsavedChanges())
        {
            return;
        }

        _isLoadingDocument = true;
        try
        {
            EditorBox.Document = new FlowDocument();
        }
        finally
        {
            _isLoadingDocument = false;
        }

        SetCurrentFile(null, DocumentType.RichText);
        MarkClean();
        UpdateFormattingControls();
        ApplyTheme();
    }

    private void RecentDocumentMenuItem_Click(object? sender, RoutedEventArgs e)
    {
        if (sender is MenuItem { Tag: string path })
        {
            OpenDocument(path, confirmUnsavedChanges: true);
        }
    }

    private bool OpenDocument(string path, bool confirmUnsavedChanges)
    {
        if (confirmUnsavedChanges && !ConfirmDiscardUnsavedChanges())
        {
            return false;
        }

        if (!File.Exists(path))
        {
            MessageBox.Show(this,
                $"The file could not be found:\n{path}",
                "File Not Found",
                MessageBoxButton.OK,
                MessageBoxImage.Warning);
            RemoveRecentDocument(path);
            return false;
        }

        LoadDocumentFromFile(path);
        return true;
    }

    private void OpenDocumentFromDialog()
    {
        if (!ConfirmDiscardUnsavedChanges())
        {
            return;
        }

        var dialog = new OpenFileDialog
        {
            Filter = "Text Documents (*.txt)|*.txt|Rich Text Format (*.rtf)|*.rtf|Markdown Files (*.md)|*.md|All files (*.*)|*.*"
        };

        if (!string.IsNullOrEmpty(_currentFilePath))
        {
            var directory = Path.GetDirectoryName(_currentFilePath);
            if (!string.IsNullOrEmpty(directory) && Directory.Exists(directory))
            {
                dialog.InitialDirectory = directory;
            }
        }

        if (dialog.ShowDialog() == true)
        {
            OpenDocument(dialog.FileName, confirmUnsavedChanges: false);
        }
    }

    private void ExitApplication()
    {
        if (!ConfirmDiscardUnsavedChanges())
        {
            return;
        }

        _allowCloseWithoutPrompt = true;
        Close();
    }

    private static string GetExtensionForDocumentType(DocumentType type)
    {
        return type switch
        {
            DocumentType.RichText => ".rtf",
            DocumentType.Markdown => ".md",
            _ => ".txt"
        };
    }

    private static int GetFilterIndexForDocumentType(DocumentType type)
    {
        return type switch
        {
            DocumentType.PlainText => 1,
            DocumentType.RichText => 2,
            DocumentType.Markdown => 3,
            _ => 1
        };
    }

    private void LoadDocumentFromFile(string path)
    {
        try
        {
            _isLoadingDocument = true;
            var documentType = DetermineDocumentType(path);
            if (documentType == DocumentType.Markdown)
            {
                _currentDocumentType = DocumentType.Markdown;
                var text = File.ReadAllText(path);
                SetEditorPlainText(text);
                SetMarkdownMode(true);
            }
            else if (documentType == DocumentType.RichText)
            {
                _currentDocumentType = DocumentType.RichText;
                LoadRtfIntoEditor(path);
                SetMarkdownMode(false);
            }
            else
            {
                _currentDocumentType = DocumentType.PlainText;
                var text = File.ReadAllText(path);
                SetEditorPlainText(text);
                SetMarkdownMode(false);
            }

            SetCurrentFile(path, _currentDocumentType);
            UpdateFormattingControls();
            MarkClean();
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Failed to load document: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
        }
        finally
        {
            _isLoadingDocument = false;
        }
    }

    private bool SaveDocumentToFile(string path)
    {
        try
        {
            var documentType = DetermineDocumentType(path);
            if (documentType == DocumentType.RichText)
            {
                if (EditorBox != null)
                {
                    var range = new TextRange(EditorBox.Document.ContentStart, EditorBox.Document.ContentEnd);
                    using var stream = new FileStream(path, FileMode.Create, FileAccess.Write);
                    range.Save(stream, DataFormats.Rtf);
                }
                SetMarkdownMode(false);
            }
            else if (documentType == DocumentType.Markdown)
            {
                var text = GetEditorPlainText();
                File.WriteAllText(path, text);
                SetMarkdownMode(true);
            }
            else
            {
                var text = GetEditorPlainText();
                File.WriteAllText(path, text);
                SetMarkdownMode(false);
            }

            SetCurrentFile(path, documentType);
            UpdateFormattingControls();
            MarkClean();
            return true;
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Failed to save document: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            return false;
        }
    }

    private static DocumentType DetermineDocumentType(string path)
    {
        var extension = Path.GetExtension(path)?.ToLowerInvariant();
        return extension switch
        {
            ".rtf" => DocumentType.RichText,
            ".md" => DocumentType.Markdown,
            _ => DocumentType.PlainText
        };
    }

    private bool SaveCurrentDocument()
    {
        if (string.IsNullOrEmpty(_currentFilePath))
        {
            return SaveDocumentWithDialog();
        }

        return SaveDocumentToFile(_currentFilePath);
    }

    private bool SaveDocumentWithDialog()
    {
        var dialog = new SaveFileDialog
        {
            Filter = "Text Documents (*.txt)|*.txt|Rich Text Format (*.rtf)|*.rtf|Markdown Files (*.md)|*.md|All files (*.*)|*.*",
            AddExtension = true,
            DefaultExt = GetExtensionForDocumentType(_currentDocumentType),
            FilterIndex = GetFilterIndexForDocumentType(_currentDocumentType)
        };

        if (!string.IsNullOrEmpty(_currentFilePath))
        {
            var directory = Path.GetDirectoryName(_currentFilePath);
            if (!string.IsNullOrEmpty(directory) && Directory.Exists(directory))
            {
                dialog.InitialDirectory = directory;
            }

            dialog.FileName = Path.GetFileName(_currentFilePath);
        }

        if (dialog.ShowDialog() == true)
        {
            return SaveDocumentToFile(dialog.FileName);
        }

        return false;
    }

    private bool ConfirmDiscardUnsavedChanges()
    {
        if (!_isDirty)
        {
            return true;
        }

        var result = MessageBox.Show(
            "You have unsaved changes. Do you want to save them before continuing?",
            "Unsaved Changes",
            MessageBoxButton.YesNoCancel,
            MessageBoxImage.Warning);

        return result switch
        {
            MessageBoxResult.Yes => SaveCurrentDocument(),
            MessageBoxResult.No => true,
            _ => false
        };
    }

    private void EditorBox_TextChanged(object sender, TextChangedEventArgs e)
    {
        if (_isLoadingDocument)
        {
            return;
        }

        MarkDirty();
        ScheduleMarkdownPreviewUpdate();
    }

    private void MainWindow_Closing(object? sender, CancelEventArgs e)
    {
        if (_allowCloseWithoutPrompt)
        {
            _allowCloseWithoutPrompt = false;
            return;
        }

        if (!ConfirmDiscardUnsavedChanges())
        {
            e.Cancel = true;
        }
    }

    protected override void OnClosed(EventArgs e)
    {
        InputLanguageManager.Current.InputLanguageChanged -= InputLanguageManager_InputLanguageChanged;
        base.OnClosed(e);
    }

    private void EditorCommand_CanExecute(object sender, CanExecuteRoutedEventArgs e)
    {
        if (EditorBox is null)
        {
            e.CanExecute = false;
            return;
        }

        e.CanExecute = e.Command switch
        {
            var command when command == ApplicationCommands.Undo => EditorBox.CanUndo,
            var command when command == ApplicationCommands.Redo => EditorBox.CanRedo,
            var command when command == ApplicationCommands.Cut => !EditorBox.Selection.IsEmpty,
            var command when command == ApplicationCommands.Copy => !EditorBox.Selection.IsEmpty,
            var command when command == ApplicationCommands.Paste => true,
            var command when command == ApplicationCommands.SelectAll => true,
            _ => false
        };

        e.Handled = true;
    }

    private void EditorCommand_Executed(object sender, ExecutedRoutedEventArgs e)
    {
        if (EditorBox is null)
        {
            return;
        }

        if (e.Command == ApplicationCommands.Undo)
        {
            EditorBox.Undo();
        }
        else if (e.Command == ApplicationCommands.Redo)
        {
            EditorBox.Redo();
        }
        else if (e.Command == ApplicationCommands.Cut)
        {
            EditorBox.Cut();
        }
        else if (e.Command == ApplicationCommands.Copy)
        {
            EditorBox.Copy();
        }
        else if (e.Command == ApplicationCommands.Paste)
        {
            EditorBox.Paste();
        }
        else if (e.Command == ApplicationCommands.SelectAll)
        {
            EditorBox.SelectAll();
        }

        e.Handled = true;
    }

    private void EditorBox_SelectionChanged(object sender, RoutedEventArgs e)
    {
        UpdateFontControlsFromSelection();
    }

    private void PopulateFontControls()
    {
        if (FontFamilyComboBox != null)
        {
            FontFamilyComboBox.ItemsSource = Fonts.SystemFontFamilies
                .OrderBy(f => f.Source)
                .ToList();
        }

        if (FontSizeComboBox != null)
        {
            FontSizeComboBox.ItemsSource = _defaultFontSizes;
        }

        UpdateFontControlsFromSelection();
    }

    private void UpdateFontControlsFromSelection()
    {
        if (FontFamilyComboBox == null || FontSizeComboBox == null)
            return;

        _isUpdatingFontControls = true;
        try
        {
            bool canFormat = CanFormat();
            FontFamilyComboBox.IsEnabled = canFormat;
            FontSizeComboBox.IsEnabled = canFormat;

            if (!canFormat || EditorBox == null)
            {
                FontFamilyComboBox.SelectedItem = null;
                FontFamilyComboBox.Text = string.Empty;
                FontSizeComboBox.SelectedItem = null;
                FontSizeComboBox.Text = string.Empty;
                return;
            }

            var selection = EditorBox.Selection;
            var familyValue = selection.GetPropertyValue(Inline.FontFamilyProperty);
            if (familyValue is FontFamily family)
            {
                var match = FontFamilyComboBox.Items.Cast<FontFamily>()
                    .FirstOrDefault(f => string.Equals(f.Source, family.Source, StringComparison.OrdinalIgnoreCase));
                FontFamilyComboBox.SelectedItem = match ?? family;
                FontFamilyComboBox.Text = family.Source;
            }
            else
            {
                FontFamilyComboBox.SelectedItem = null;
                FontFamilyComboBox.Text = string.Empty;
            }

            var sizeValue = selection.GetPropertyValue(Inline.FontSizeProperty);
            if (sizeValue is double size)
            {
                FontSizeComboBox.Text = size.ToString("0.#");
                var closest = _defaultFontSizes.FirstOrDefault(s => Math.Abs(s - size) < 0.1);
                FontSizeComboBox.SelectedItem = closest > 0 ? closest : null;
            }
            else
            {
                FontSizeComboBox.SelectedItem = null;
                FontSizeComboBox.Text = string.Empty;
            }
        }
        finally
        {
            _isUpdatingFontControls = false;
        }
    }

    private void ApplyFontFamily(FontFamily fontFamily)
    {
        if (EditorBox == null || !CanFormat())
            return;

        EditorBox.Selection.ApplyPropertyValue(Inline.FontFamilyProperty, fontFamily);
    }

    private void ApplyFontSize(double fontSize)
    {
        if (EditorBox == null || !CanFormat())
            return;

        double clampedSize = Math.Min(MaxFontSize, Math.Max(MinFontSize, fontSize));
        EditorBox.Selection.ApplyPropertyValue(Inline.FontSizeProperty, clampedSize);
    }

    private void FontFamilyComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (_isUpdatingFontControls || !CanFormat())
            return;

        if (FontFamilyComboBox?.SelectedItem is FontFamily family)
        {
            ApplyFontFamily(family);
            MarkDirty();
        }
    }

    private void FontSizeComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (_isUpdatingFontControls || !CanFormat())
            return;

        if (FontSizeComboBox?.SelectedItem is double size)
        {
            ApplyFontSize(size);
            MarkDirty();
        }
    }

    private void FontSizeComboBox_LostFocus(object sender, RoutedEventArgs e)
    {
        ApplyFontSizeFromTextInput();
    }

    private void FontSizeComboBox_PreviewKeyDown(object sender, KeyEventArgs e)
    {
        if (e.Key == Key.Enter)
        {
            ApplyFontSizeFromTextInput();
            e.Handled = true;
        }
    }

    private void ApplyFontSizeFromTextInput()
    {
        if (_isUpdatingFontControls || !CanFormat() || FontSizeComboBox == null)
            return;

        if (double.TryParse(FontSizeComboBox.Text, out double size))
        {
            ApplyFontSize(size);
            FontSizeComboBox.Text = Math.Min(MaxFontSize, Math.Max(MinFontSize, size)).ToString("0.#");
            MarkDirty();
        }
        else
        {
            UpdateFontControlsFromSelection();
        }
    }

    private void ToggleBold()
    {
        if (EditorBox == null || !CanFormat())
            return;

        TextSelection selection = EditorBox.Selection;
        object currentWeight = selection.GetPropertyValue(Inline.FontWeightProperty);

        if (currentWeight != DependencyProperty.UnsetValue && currentWeight.Equals(FontWeights.Bold))
        {
            selection.ApplyPropertyValue(Inline.FontWeightProperty, FontWeights.Normal);
        }
        else
        {
            selection.ApplyPropertyValue(Inline.FontWeightProperty, FontWeights.Bold);
        }
    }

    private void ToggleUnderline()
    {
        if (EditorBox == null || !CanFormat())
            return;

        TextSelection selection = EditorBox.Selection;
        object currentDecorations = selection.GetPropertyValue(Inline.TextDecorationsProperty);

        var current = currentDecorations as TextDecorationCollection;
        bool isUnderlined = current != null && current.Contains(TextDecorations.Underline[0]);

        if (isUnderlined)
        {
            // Remove underline
            selection.ApplyPropertyValue(Inline.TextDecorationsProperty, null);
        }
        else
        {
            // Apply underline
            selection.ApplyPropertyValue(Inline.TextDecorationsProperty, TextDecorations.Underline);
        }
    }

    private void BoldButton_Click(object sender, RoutedEventArgs e)
    {
        if (!CanFormat())
            return;

        ToggleBold();
        MarkDirty();
    }

    private void UnderlineButton_Click(object sender, RoutedEventArgs e)
    {
        if (!CanFormat())
            return;

        ToggleUnderline();
        MarkDirty();
    }

    private bool CanFormat()
    {
        return _currentDocumentType == DocumentType.RichText;
    }

    private void UpdateFormattingControls()
    {
        bool canFormat = CanFormat();

        if (BoldButton != null)
            BoldButton.IsEnabled = canFormat;
        if (UnderlineButton != null)
            UnderlineButton.IsEnabled = canFormat;

        UpdateFontControlsFromSelection();
    }
}

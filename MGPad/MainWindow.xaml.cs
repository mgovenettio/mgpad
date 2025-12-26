using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Threading;
using System.Xml.Linq;
using Markdig;
using PdfSharp.Drawing;
using PdfSharp.Pdf;
using DrawingColor = System.Drawing.Color;
using FontFamily = System.Windows.Media.FontFamily;
using KeyEventArgs = System.Windows.Input.KeyEventArgs;
using Color = System.Windows.Media.Color;
using Brush = System.Windows.Media.Brush;
using WpfTextBox = System.Windows.Controls.TextBox;

namespace MGPad;

public enum DocumentType
{
    RichText,
    PlainText,
    Markdown,
    OpenDocument
}

internal readonly record struct ParsedNumber(decimal Value, int FractionDigits);

public partial class MainWindow : Window
{
    public static readonly RoutedUICommand ToggleBoldCommand =
        new RoutedUICommand("ToggleBold", "ToggleBold", typeof(MainWindow),
            new InputGestureCollection { new KeyGesture(Key.B, ModifierKeys.Control) });

    public static readonly RoutedUICommand ToggleItalicCommand =
        new RoutedUICommand("ToggleItalic", "ToggleItalic", typeof(MainWindow),
            new InputGestureCollection { new KeyGesture(Key.I, ModifierKeys.Control) });

    public static readonly RoutedUICommand ToggleUnderlineCommand =
        new RoutedUICommand("ToggleUnderline", "ToggleUnderline", typeof(MainWindow),
            new InputGestureCollection { new KeyGesture(Key.U, ModifierKeys.Control) });

    public static readonly RoutedUICommand ToggleMonospacedCommand =
        new RoutedUICommand("ToggleMonospaced", "ToggleMonospaced", typeof(MainWindow),
            new InputGestureCollection { new KeyGesture(Key.M, ModifierKeys.Control | ModifierKeys.Shift) });

    public static readonly RoutedUICommand ToggleStrikethroughCommand =
        new RoutedUICommand("ToggleStrikethrough", "ToggleStrikethrough", typeof(MainWindow),
            new InputGestureCollection { new KeyGesture(Key.X, ModifierKeys.Control | ModifierKeys.Shift) });

    public static readonly RoutedUICommand ToggleInputLanguageCommand =
        new RoutedUICommand("ToggleInputLanguage", "ToggleInputLanguage", typeof(MainWindow),
            new InputGestureCollection { new KeyGesture(Key.J, ModifierKeys.Control) });

    public static readonly RoutedUICommand InsertTimestampCommand =
        new RoutedUICommand("InsertTimestamp", "InsertTimestamp", typeof(MainWindow),
            new InputGestureCollection { new KeyGesture(Key.T, ModifierKeys.Alt | ModifierKeys.Shift) });

    public static readonly DependencyProperty IsSpellCheckEnabledProperty =
        DependencyProperty.Register(nameof(IsSpellCheckEnabled), typeof(bool), typeof(MainWindow),
            new PropertyMetadata(true, OnIsSpellCheckEnabledChanged));

    private string? _currentFilePath;
    private DocumentType _currentDocumentType;
    private DocumentType _previousDocumentType = DocumentType.RichText;
    private bool _isDirty;
    private bool _isLoadingDocument;
    private bool _allowCloseWithoutPrompt;
    private bool _isMarkdownMode = false;
    private bool _isNightMode = false;
    private const double ZoomStep = 0.1;
    private const double MinZoom = 0.5;
    private const double MaxZoom = 3.0;
    private const double DefaultZoom = 1.0;
    private double _zoomLevel = DefaultZoom;
    private double _lastAppliedZoom = 1.0;
    private const double MinFontSize = 6;
    private const double MaxFontSize = 96;
    private const int MaxRecentDocuments = 10;
    private readonly List<string> _recentDocuments = new();
    private readonly string _recentDocumentsFilePath;
    private readonly string _settingsFilePath;
    private readonly string _customDictionaryPath;
    private readonly DispatcherTimer _markdownPreviewTimer;
    private CultureInfo? _englishInputLanguage;
    private CultureInfo? _japaneseInputLanguage;
    private bool _isUpdatingFontControls;
    private bool _isLoadingSettings;
    private readonly double[] _defaultFontSizes = new double[]
        { 8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72 };
    private WpfTextBox? _fontSizeEditableTextBox;
    private FontFamily? _lastKnownFontFamily;
    private double? _lastKnownFontSize;
    private static double PointsToDips(double points) => points * 96.0 / 72.0;

    private static double DipsToPoints(double dips) => dips * 72.0 / 96.0;

    private static double ClampFontSizeInPoints(double fontSize) =>
        Math.Min(MaxFontSize, Math.Max(MinFontSize, fontSize));
    private static readonly Regex NumberRegex = new(
        "[-+]?(?:\\d+\\.?\\d*|\\.\\d+)(?:[eE][-+]?\\d+)?",
        RegexOptions.Compiled);
    private static readonly Regex WordRegex = new(
        "\\b[^\\s]+\\b",
        RegexOptions.Compiled);
    private readonly MarkdownPipeline _markdownPipeline;
    private readonly AutosaveService _autosaveService;
    private Guid _untitledDocumentId = Guid.NewGuid();
    private bool _isRecoveredDocument;
    private readonly List<string> _pendingRecoveryAutosavePaths = new();
    // Keep list indentation uniform across bullet, numbered, and lettered markers so the
    // content column lines up regardless of the marker glyph width. The marker offset is
    // kept inside the list bounds so bullets are visible within the editor border padding.
    private static readonly Thickness ListIndentationMargin = new(26, 0, 0, 0);
    private const double ListIndentationMarkerOffset = 10;
    private static readonly Thickness ListIndentationPadding = new(0);
    private static readonly object SpellCheckContextMenuTag = new();
    private static readonly TimeSpan AutosaveInterval = TimeSpan.FromSeconds(60);

    private sealed class StyleConfiguration
    {
        // Body and mono font stacks prefer Windows defaults while including explicit Japanese
        // fallbacks.  The comma-separated list follows the same syntax as CSS font-family:
        // choose the first installed face and fall back to later entries when characters (e.g.,
        // CJK glyphs) are missing.  This keeps Unicode-heavy notes readable without requiring
        // the exact font set from the authoring machine.
        public string BodyFontFamily { get; init; } = "Times New Roman, 'Yu Gothic UI'";

        // Monospaced stack favors Cascadia Code/Cascadia Mono, then Consolas, and finally MS
        // Gothic so code snippets remain aligned even on systems that only have Japanese-centric
        // fonts available.
        public string MonoFontFamily { get; init; } = "Cascadia Code, Consolas, 'MS Gothic'";
    }

    private readonly StyleConfiguration _styleConfiguration = new();

    private FontFamily GetBodyFontFamily() => new(_styleConfiguration.BodyFontFamily);

    private FontFamily GetMonoFontFamily() => new(_styleConfiguration.MonoFontFamily);

    private sealed class PdfTextRun
    {
        public string Text { get; set; } = string.Empty;
        public bool IsBold { get; set; }
        public bool IsItalic { get; set; }
        public bool IsUnderline { get; set; }
        public bool IsStrikethrough { get; set; }

        // Flag propagated through the UI toggle, PDF export, and ODT serialization paths to mark
        // code-like spans.  When true, the run always uses the configured monospaced stack instead
        // of inheriting the surrounding paragraph font.
        public bool IsMonospaced { get; set; }
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

    private sealed class UserSettings
    {
        public bool IsSpellCheckEnabled { get; set; } = true;
    }

    private const string HelpText =
        "MGPad quick help:\n" +
        "\n" +
        "• Ctrl+N / Ctrl+O / Ctrl+S to create, open, or save documents.\n" +
        "• Ctrl+B / Ctrl+I / Ctrl+U / Ctrl+Shift+X to toggle bold, italic, underline, or strikethrough.\n" +
        "• Alt+Shift+T inserts a timestamp.\n" +
        "• Use View → Markdown Mode to preview Markdown on the right.\n" +
        "• Toggle JP/EN switches input language when both are available.\n" +
        "• Use the zoom controls in the status bar to adjust text size.";

    public bool IsSpellCheckEnabled
    {
        get => (bool)GetValue(IsSpellCheckEnabledProperty);
        set => SetValue(IsSpellCheckEnabledProperty, value);
    }

    public MainWindow(RecoveryItem? recoveryItem = null)
    {
        InitializeComponent();
        _recentDocumentsFilePath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "MGPad",
            "recent-documents.txt");
        _settingsFilePath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "MGPad",
            "settings.json");
        _customDictionaryPath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "MGPad",
            "custom-dictionary.txt");
        LoadUserSettings();
        ConfigureCustomDictionary();
        _markdownPipeline = new MarkdownPipelineBuilder()
            .UseAdvancedExtensions()
            .DisableHtml()
            .Build();
        _autosaveService = new AutosaveService(
            AutosaveInterval,
            () => _isDirty,
            GetAutosaveContext,
            SaveAutosaveDocument);
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

        CommandBindings.Add(new CommandBinding(ToggleItalicCommand,
            (s, e) =>
            {
                if (!CanFormat())
                {
                    e.Handled = true;
                    return;
                }

                ToggleItalic();
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
        CommandBindings.Add(new CommandBinding(ToggleStrikethroughCommand,
            (s, e) =>
            {
                if (!CanFormat())
                {
                    e.Handled = true;
                    return;
                }

                ToggleStrikethrough();
                MarkDirty();
                e.Handled = true;
            }));
        CommandBindings.Add(new CommandBinding(
            ToggleMonospacedCommand,
            (s, e) =>
            {
                if (!CanFormat())
                {
                    e.Handled = true;
                    return;
                }

                ToggleMonospacedFromUi();
                e.Handled = true;
            },
            (s, e) => e.CanExecute = CanFormat()));
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
            EditorBox.PreviewKeyDown += EditorBox_PreviewKeyDown;
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
        ApplyDefaultStyleToCurrentDocument();
        UpdateFormattingControls();
        UpdateLanguageIndicator();

        SetMarkdownMode(false);
        ApplyTheme();
        ApplyZoom();
        LoadRecentDocuments();
        UpdateRecentDocumentsMenu();
        _autosaveService.Start();

        if (recoveryItem != null)
        {
            LoadRecoveredDocument(recoveryItem);
        }
    }

    private void HelpMenuItem_Click(object sender, RoutedEventArgs e)
    {
        System.Windows.MessageBox.Show(
            HelpText,
            "MGPad Help",
            MessageBoxButton.OK,
            MessageBoxImage.None);
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
        bool isMarkdownDocument = _currentDocumentType == DocumentType.Markdown
            || (!string.IsNullOrEmpty(_currentFilePath) && DetermineDocumentType(_currentFilePath) == DocumentType.Markdown);

        if (enable)
        {
            if (!isMarkdownDocument)
                _previousDocumentType = _currentDocumentType;

            if (_currentDocumentType != DocumentType.Markdown)
                _currentDocumentType = DocumentType.Markdown;
        }
        else if (!isMarkdownDocument)
        {
            _currentDocumentType = _previousDocumentType;
        }

        _isMarkdownMode = enable;

        if (MarkdownModeMenuItem != null)
            MarkdownModeMenuItem.IsChecked = enable;

        UpdateFormattingControls();
        ApplyMarkdownModeLayout();
        UpdateMarkdownPreview();
    }

    private void DisableMarkdownModeLayout()
    {
        _isMarkdownMode = false;

        if (MarkdownModeMenuItem != null)
            MarkdownModeMenuItem.IsChecked = false;

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
            : System.Windows.Media.Brushes.White;
        Brush panelBackground = _isNightMode
            ? new SolidColorBrush(Color.FromRgb(45, 45, 48))
            : System.Windows.Media.Brushes.WhiteSmoke;
        Brush borderBrush = _isNightMode
            ? new SolidColorBrush(Color.FromRgb(62, 62, 66))
            : System.Windows.Media.Brushes.LightGray;
        Brush foreground = _isNightMode ? System.Windows.Media.Brushes.WhiteSmoke : System.Windows.Media.Brushes.Black;

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

        if (MainStatusBar != null)
        {
            MainStatusBar.Background = panelBackground;
            MainStatusBar.Foreground = foreground;
        }

        if (ZoomPercentageTextBlock != null)
            ZoomPercentageTextBlock.Foreground = foreground;

        if (NightModeButton != null)
            NightModeButton.Content = _isNightMode ? "Day" : "Night";

        UpdateMarkdownPreview();
    }

    private void ResetZoomStateForDocument()
    {
        _lastAppliedZoom = 1.0;
        if (EditorBox != null)
            EditorBox.LayoutTransform = Transform.Identity;
    }

    private static void ScaleInlineFontSize(Inline inline, double scaleRatio)
    {
        object fontSizeValue = inline.ReadLocalValue(TextElement.FontSizeProperty);
        if (fontSizeValue is double inlineSize)
            inline.SetValue(TextElement.FontSizeProperty, inlineSize * scaleRatio);

        if (inline is Span span)
        {
            foreach (Inline child in span.Inlines)
            {
                ScaleInlineFontSize(child, scaleRatio);
            }
        }
    }

    private void ScaleTextElementFontSize(TextElement element, double scaleRatio)
    {
        object fontSizeValue = element.ReadLocalValue(TextElement.FontSizeProperty);
        if (fontSizeValue is double elementSize)
            element.SetValue(TextElement.FontSizeProperty, elementSize * scaleRatio);

        switch (element)
        {
            case Paragraph paragraph:
                foreach (Inline inline in paragraph.Inlines)
                {
                    ScaleInlineFontSize(inline, scaleRatio);
                }
                break;
            case Section section:
                foreach (Block child in section.Blocks)
                {
                    ScaleTextElementFontSize(child, scaleRatio);
                }
                break;
            case List list:
                foreach (ListItem item in list.ListItems)
                {
                    foreach (Block child in item.Blocks)
                    {
                        ScaleTextElementFontSize(child, scaleRatio);
                    }
                }
                break;
            case Table table:
                foreach (TableRowGroup group in table.RowGroups)
                {
                    foreach (TableRow row in group.Rows)
                    {
                        foreach (TableCell cell in row.Cells)
                        {
                            foreach (Block child in cell.Blocks)
                            {
                                ScaleTextElementFontSize(child, scaleRatio);
                            }
                        }
                    }
                }
                break;
            case Figure figure:
                foreach (Block child in figure.Blocks)
                {
                    ScaleTextElementFontSize(child, scaleRatio);
                }
                break;
            case Floater floater:
                foreach (Block child in floater.Blocks)
                {
                    ScaleTextElementFontSize(child, scaleRatio);
                }
                break;
        }
    }

    private static Thickness ScaleThickness(Thickness thickness, double scaleRatio)
    {
        return new Thickness(
            thickness.Left * scaleRatio,
            thickness.Top * scaleRatio,
            thickness.Right * scaleRatio,
            thickness.Bottom * scaleRatio);
    }

    private void ScaleDocumentLayout(FlowDocument document, double scaleRatio)
    {
        document.FontSize *= scaleRatio;

        if (!double.IsNaN(document.PageWidth) && !double.IsInfinity(document.PageWidth))
        {
            document.PageWidth *= scaleRatio;
        }

        document.PagePadding = ScaleThickness(document.PagePadding, scaleRatio);

        foreach (Block block in document.Blocks)
        {
            ScaleTextElementFontSize(block, scaleRatio);
        }
    }

    private void ApplyZoom()
    {
        if (EditorBox != null)
            EditorBox.LayoutTransform = Transform.Identity;

        if (EditorBox?.Document != null)
        {
            double scaleRatio = _zoomLevel / _lastAppliedZoom;
            if (Math.Abs(scaleRatio - 1.0) > 0.0001)
            {
                ScaleDocumentLayout(EditorBox.Document, scaleRatio);
                _lastAppliedZoom = _zoomLevel;
            }
        }

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
        if (!_isMarkdownMode || MarkdownPreviewBrowser == null)
            return;

        string markdown = GetEditorPlainText();
        string normalized = markdown.Replace("\r\n", "\n").Replace("\r", "\n");

        string html = string.IsNullOrWhiteSpace(normalized)
            ? string.Empty
            : Markdown.ToHtml(normalized, _markdownPipeline);

        string background = _isNightMode ? "#2d2d30" : "#ffffff";
        string foreground = _isNightMode ? "#f5f5f5" : "#000000";

        string document = $"<!DOCTYPE html><html><head><meta charset=\"utf-8\" /><style>body {{ font-family: {_styleConfiguration.BodyFontFamily}, serif; padding: 12px; color: {foreground}; background: {background}; }} code, pre {{ font-family: 'Cascadia Code', Consolas, 'Courier New', monospace; }} pre {{ background: rgba(0,0,0,0.04); padding: 8px; overflow-x: auto; }} a {{ color: #0066cc; }}</style></head><body>{html}</body></html>";

        MarkdownPreviewBrowser.NavigateToString(document);
    }

    private FlowDocument CreateStyledDocument()
    {
        return new FlowDocument
        {
            FontFamily = GetBodyFontFamily(),
            FontSize = PointsToDips(12)
        };
    }

    private void ApplyDefaultStyleToCurrentDocument()
    {
        if (EditorBox?.Document != null)
        {
            EditorBox.Document.FontFamily = GetBodyFontFamily();
        }
    }

    private bool IsMonospacedFont(FontFamily family)
    {
        return string.Equals(
            family.Source,
            GetMonoFontFamily().Source,
            StringComparison.OrdinalIgnoreCase);
    }

    private void SetEditorPlainText(string text)
    {
        if (EditorBox == null)
            return;

        EditorBox.Document = CreateStyledDocument();
        var range = new TextRange(EditorBox.Document.ContentStart, EditorBox.Document.ContentEnd);
        range.Text = text;
        ApplyTheme();
        ResetZoomStateForDocument();
        ApplyZoom();
    }

    private void LoadRtfIntoEditor(string path)
    {
        if (EditorBox == null)
            return;

        EditorBox.Document = CreateStyledDocument();
        var range = new TextRange(EditorBox.Document.ContentStart, EditorBox.Document.ContentEnd);

        using var stream = new FileStream(path, FileMode.Open, FileAccess.Read);
        range.Load(stream, System.Windows.DataFormats.Rtf);
        ApplyTheme();
        ResetZoomStateForDocument();
        ApplyZoom();
    }

    private sealed record OdtTextStyle(
        bool Bold,
        bool Italic,
        bool Underline,
        string? FontFamily,
        bool IsMonospaced);

    private HashSet<string> GetKnownMonospacedFamilies()
    {
        var result = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "Cascadia Code",
            "Cascadia Mono",
            "Consolas",
            "Courier New",
            "JetBrains Mono",
            "Liberation Mono",
            "Lucida Console",
            "Menlo",
            "Monaco",
            "Roboto Mono",
            "Source Code Pro"
        };

        foreach (var family in _styleConfiguration.MonoFontFamily.Split(','))
        {
            string cleaned = family.Trim().Trim('\'', '"');
            if (!string.IsNullOrWhiteSpace(cleaned))
            {
                result.Add(cleaned);
            }
        }

        return result;
    }

    private void LoadOdtIntoEditor(string path)
    {
        if (EditorBox == null)
            return;

        try
        {
            var paragraphs = ParseOdtDocument(path);
            var document = CreateStyledDocument();

            foreach (var paragraph in paragraphs)
            {
                var wpfParagraph = new Paragraph();

                foreach (var run in paragraph.Runs)
                {
                    AppendRunToParagraph(wpfParagraph, run);
                }

                document.Blocks.Add(wpfParagraph);
            }

            EditorBox.Document = document;
            ApplyTheme();
            ResetZoomStateForDocument();
            ApplyZoom();
        }
        catch
        {
            var text = File.ReadAllText(path);
            SetEditorPlainText(text);
        }
    }

    private void AppendRunToParagraph(Paragraph paragraph, PdfTextRun run)
    {
        string[] segments = run.Text.Replace("\r\n", "\n").Split('\n');
        for (int i = 0; i < segments.Length; i++)
        {
            if (segments[i].Length > 0)
            {
                var wpfRun = new Run(segments[i]);

                if (run.IsBold)
                    wpfRun.FontWeight = FontWeights.Bold;
                if (run.IsItalic)
                    wpfRun.FontStyle = System.Windows.FontStyles.Italic;
                if (run.IsUnderline)
                    wpfRun.TextDecorations = TextDecorations.Underline;
                if (run.IsStrikethrough)
                {
                    var decorations = wpfRun.TextDecorations ?? new TextDecorationCollection();
                    decorations.Add(TextDecorations.Strikethrough[0]);
                    wpfRun.TextDecorations = decorations;
                }
                if (run.IsMonospaced)
                    wpfRun.FontFamily = GetMonoFontFamily();

                paragraph.Inlines.Add(wpfRun);
            }

            if (i < segments.Length - 1)
            {
                paragraph.Inlines.Add(new LineBreak());
            }
        }
    }

    private List<PdfParagraph> ParseOdtDocument(string path)
    {
        using var archive = ZipFile.OpenRead(path);
        var contentEntry = archive.GetEntry("content.xml") ?? throw new InvalidOperationException("content.xml missing");

        XDocument contentDoc = LoadXmlFromEntry(contentEntry);
        XDocument? stylesDoc = archive.GetEntry("styles.xml") is { } stylesEntry
            ? LoadXmlFromEntry(stylesEntry)
            : null;

        // ODT stores formatting via named text styles and optional font-face definitions.  We
        // collect every known style from content.xml and styles.xml, augment it with a curated list
        // of monospaced font families (to avoid missing glyphs on non-English systems), and then
        // resolve each span by merging inherited styles as we walk the tree.
        var monospacedFamilies = GetKnownMonospacedFamilies();
        var styles = new Dictionary<string, OdtTextStyle>(StringComparer.OrdinalIgnoreCase);

        MergeOdtStyles(styles, contentDoc, monospacedFamilies);
        if (stylesDoc != null)
        {
            MergeOdtStyles(styles, stylesDoc, monospacedFamilies);
        }

        var paragraphs = new List<PdfParagraph>();
        XNamespace text = contentDoc.Root?.GetNamespaceOfPrefix("text") ?? "text";

        foreach (var paragraph in contentDoc.Descendants(text + "p"))
        {
            var pdfParagraph = new PdfParagraph();

            foreach (var node in paragraph.Nodes())
            {
                AppendRunsFromOdtNode(node, pdfParagraph.Runs, styles, text, null);
            }

            if (pdfParagraph.Runs.Count > 0)
            {
                paragraphs.Add(pdfParagraph);
            }
        }

        return paragraphs;
    }

    private static XDocument LoadXmlFromEntry(ZipArchiveEntry entry)
    {
        using var stream = entry.Open();
        using var reader = new StreamReader(stream, Encoding.UTF8);
        return XDocument.Load(reader);
    }

    private void MergeOdtStyles(
        Dictionary<string, OdtTextStyle> accumulator,
        XDocument document,
        HashSet<string> monospacedFamilies)
    {
        var fontFaces = ParseFontFaces(document);
        foreach (var (name, style) in ParseTextStyles(document, fontFaces, monospacedFamilies))
        {
            accumulator[name] = style;
        }
    }

    private Dictionary<string, string> ParseFontFaces(XDocument document)
    {
        var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        XNamespace style = document.Root?.GetNamespaceOfPrefix("style") ?? "style";
        XNamespace svg = document.Root?.GetNamespaceOfPrefix("svg") ?? "svg";

        foreach (var fontFace in document.Descendants(style + "font-face"))
        {
            string? name = (string?)fontFace.Attribute(style + "name");
            string? family = (string?)fontFace.Attribute(svg + "font-family")
                ?? (string?)fontFace.Attribute(style + "font-family-generic");

            if (!string.IsNullOrEmpty(name) && !string.IsNullOrEmpty(family))
            {
                result[name] = family.Trim('\"', '\'');
            }
        }

        return result;
    }

    private Dictionary<string, OdtTextStyle> ParseTextStyles(
        XDocument document,
        Dictionary<string, string> fontFaces,
        HashSet<string> monospacedFamilies)
    {
        var styles = new Dictionary<string, OdtTextStyle>(StringComparer.OrdinalIgnoreCase);

        XNamespace style = document.Root?.GetNamespaceOfPrefix("style") ?? "style";
        XNamespace fo = document.Root?.GetNamespaceOfPrefix("fo") ?? "fo";

        foreach (var styleElement in document.Descendants(style + "style"))
        {
            string? name = (string?)styleElement.Attribute(style + "name");
            string? family = (string?)styleElement.Attribute(style + "family");
            if (string.IsNullOrEmpty(name) || !string.Equals(family, "text", StringComparison.OrdinalIgnoreCase))
                continue;

            var textProps = styleElement.Element(style + "text-properties");
            bool isBold = string.Equals((string?)textProps?.Attribute(fo + "font-weight"), "bold", StringComparison.OrdinalIgnoreCase);
            bool isItalic = string.Equals((string?)textProps?.Attribute(fo + "font-style"), "italic", StringComparison.OrdinalIgnoreCase);
            bool isUnderline = textProps?.Attribute(style + "text-underline-style") != null;

            string? fontName = (string?)textProps?.Attribute(style + "font-name");
            string? fontFamily = fontName != null && fontFaces.TryGetValue(fontName, out var resolved)
                ? resolved
                : (string?)textProps?.Attribute(fo + "font-family");
            fontFamily = fontFamily?.Trim('\"', '\'');

            bool monospacedByName = string.Equals(name, "Code", StringComparison.OrdinalIgnoreCase)
                || string.Equals(name, "Mono", StringComparison.OrdinalIgnoreCase);
            bool monospacedByFont = !string.IsNullOrEmpty(fontFamily) && monospacedFamilies.Contains(fontFamily);

            styles[name] = new OdtTextStyle(
                isBold,
                isItalic,
                isUnderline,
                fontFamily,
                monospacedByName || monospacedByFont);
        }

        return styles;
    }

    private void AppendRunsFromOdtNode(
        XNode node,
        List<PdfTextRun> runs,
        Dictionary<string, OdtTextStyle> styles,
        XNamespace text,
        OdtTextStyle? currentStyle)
    {
        if (node is XText textNode)
        {
            if (!string.IsNullOrEmpty(textNode.Value))
            {
                runs.Add(CreateRunFromOdtText(textNode.Value, currentStyle));
            }
            return;
        }

        if (node is not XElement element)
            return;

        if (element.Name == text + "span")
        {
            string? styleName = (string?)element.Attribute(text + "style-name");
            styles.TryGetValue(styleName ?? string.Empty, out var styleFromSpan);
            var merged = MergeStyles(currentStyle, styleFromSpan);

            foreach (var child in element.Nodes())
            {
                AppendRunsFromOdtNode(child, runs, styles, text, merged);
            }
            return;
        }

        if (element.Name == text + "s")
        {
            int count = (int?)element.Attribute(text + "c") ?? 1;
            runs.Add(CreateRunFromOdtText(new string(' ', Math.Max(1, count)), currentStyle));
            return;
        }

        if (element.Name == text + "line-break")
        {
            runs.Add(CreateRunFromOdtText("\n", currentStyle));
            return;
        }

        foreach (var child in element.Nodes())
        {
            AppendRunsFromOdtNode(child, runs, styles, text, currentStyle);
        }
    }

    private PdfTextRun CreateRunFromOdtText(string text, OdtTextStyle? style)
    {
        return new PdfTextRun
        {
            Text = text,
            IsBold = style?.Bold ?? false,
            IsItalic = style?.Italic ?? false,
            IsUnderline = style?.Underline ?? false,
            IsStrikethrough = false,
            IsMonospaced = style?.IsMonospaced ?? false
        };
    }

    private OdtTextStyle MergeStyles(OdtTextStyle? baseStyle, OdtTextStyle? overrideStyle)
    {
        if (baseStyle == null)
            return overrideStyle ?? new OdtTextStyle(false, false, false, null, false);
        if (overrideStyle == null)
            return baseStyle;

        return new OdtTextStyle(
            baseStyle.Bold || overrideStyle.Bold,
            baseStyle.Italic || overrideStyle.Italic,
            baseStyle.Underline || overrideStyle.Underline,
            overrideStyle.FontFamily ?? baseStyle.FontFamily,
            baseStyle.IsMonospaced || overrideStyle.IsMonospaced);
    }

    private string GetEditorPlainText()
    {
        if (EditorBox == null)
            return string.Empty;

        TextRange range = new TextRange(EditorBox.Document.ContentStart, EditorBox.Document.ContentEnd);
        return range.Text ?? string.Empty;
    }

    private static string NormalizeTextForCounting(string text)
    {
        if (string.IsNullOrEmpty(text))
            return string.Empty;

        string normalized = text.Replace("\r\n", "\n").Replace("\r", "\n");
        return normalized.TrimEnd('\n');
    }

    private static (int wordCount, int charCount) CountWordsAndCharacters(string text)
    {
        if (string.IsNullOrEmpty(text))
            return (0, 0);

        int wordCount = WordRegex.Matches(text).Count;
        int charCount = text.Length;

        return (wordCount, charCount);
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

        string timestamp = DateTime.Now.ToString("h:mm tt") + ": ";

        TextRange selection = EditorBox.Selection;

        if (!selection.IsEmpty)
        {
            selection.Text = timestamp;
            selection.ApplyPropertyValue(Inline.FontWeightProperty, FontWeights.Bold);
            EditorBox.CaretPosition = selection.End;
        }
        else
        {
            var caret = EditorBox.CaretPosition;

            if (!caret.IsAtInsertionPosition)
                caret = caret.GetInsertionPosition(LogicalDirection.Forward);

            selection = new TextRange(caret, caret);
            selection.Text = timestamp;
            selection.ApplyPropertyValue(Inline.FontWeightProperty, FontWeights.Bold);

            EditorBox.CaretPosition = selection.End;
        }

        EditorBox.Focus();
        MarkDirty();
        ScheduleMarkdownPreviewUpdate();
    }

    private void SumSelectionMenuItem_Click(object sender, RoutedEventArgs e)
    {
        if (EditorBox == null)
            return;

        string selectedText = EditorBox.Selection.Text;
        if (string.IsNullOrWhiteSpace(selectedText))
        {
            System.Windows.MessageBox.Show(
                "Select one or more numbers to sum.",
                "Sum selection",
                MessageBoxButton.OK,
                MessageBoxImage.None);
            return;
        }

        var numbers = ExtractNumbersFromSelection(selectedText);
        if (numbers.Count == 0)
        {
            System.Windows.MessageBox.Show(
                "No numbers were found in the selection.",
                "Sum selection",
                MessageBoxButton.OK,
                MessageBoxImage.None);
            return;
        }

        decimal sum = numbers.Sum(number => number.Value);
        int maxFractionDigits = numbers.Count > 0 ? numbers.Max(number => number.FractionDigits) : 0;
        string sumText = sum.ToString($"F{maxFractionDigits}", CultureInfo.InvariantCulture);
        string insertionText = GetInsertionText(selectedText, sumText);

        var insertionPosition = EditorBox.Selection.End;
        if (!insertionPosition.IsAtInsertionPosition)
            insertionPosition = insertionPosition.GetInsertionPosition(LogicalDirection.Forward) ?? insertionPosition;

        var insertionRange = new TextRange(insertionPosition, insertionPosition);
        insertionRange.Text = insertionText;

        EditorBox.CaretPosition = insertionRange.End;
        EditorBox.Focus();
        MarkDirty();
        ScheduleMarkdownPreviewUpdate();
    }

    internal static List<ParsedNumber> ExtractNumbersFromSelection(string selectedText)
    {
        var numbers = new List<ParsedNumber>();

        foreach (Match match in NumberRegex.Matches(selectedText))
        {
            if (decimal.TryParse(match.Value, NumberStyles.Float, CultureInfo.InvariantCulture, out decimal value))
            {
                string matchValue = match.Value;
                int exponentIndex = matchValue.IndexOfAny(new[] { 'e', 'E' });
                string coefficient = exponentIndex >= 0 ? matchValue[..exponentIndex] : matchValue;
                int decimalPointIndex = coefficient.IndexOf('.');
                int fractionDigits = decimalPointIndex >= 0
                    ? coefficient.Length - decimalPointIndex - 1
                    : 0;

                numbers.Add(new ParsedNumber(value, fractionDigits));
            }
        }

        return numbers;
    }

    private static string GetInsertionText(string selectedText, string sumText)
    {
        bool containsNewLine = selectedText.Contains('\n') || selectedText.Contains('\r');
        if (containsNewLine)
        {
            if (selectedText.EndsWith("\n") || selectedText.EndsWith("\r"))
                return sumText;

            return Environment.NewLine + sumText;
        }

        char lastChar = selectedText.LastOrDefault();
        if (lastChar == ',')
            return " " + sumText;

        if (lastChar != default && !char.IsWhiteSpace(lastChar))
            return " " + sumText;

        return sumText;
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

    private void SetCurrentFile(string? path, DocumentType type, bool clearRecoveryState = true)
    {
        if (clearRecoveryState)
        {
            ClearRecoveredState(deletePendingFiles: true);
        }

        _currentFilePath = path;
        _currentDocumentType = type;
        if (string.IsNullOrEmpty(path))
        {
            _untitledDocumentId = Guid.NewGuid();
        }
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

    private void ClearRecoveredState(bool deletePendingFiles)
    {
        if (deletePendingFiles)
        {
            DeletePendingRecoveryAutosaveFiles();
        }

        _isRecoveredDocument = false;
    }

    private void DeletePendingRecoveryAutosaveFiles()
    {
        foreach (var path in _pendingRecoveryAutosavePaths)
        {
            AutosaveService.TryDeleteFile(path);
        }

        _pendingRecoveryAutosavePaths.Clear();
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
        var baseName = string.IsNullOrEmpty(_currentFilePath)
            ? "Untitled"
            : Path.GetFileName(_currentFilePath);

        if (_isRecoveredDocument)
        {
            baseName += " [Recovered]";
        }

        return baseName;
    }

    private AutosaveContext GetAutosaveContext()
    {
        var documentType = GetCurrentDocumentTypeForSaving();
        var extension = GetExtensionForDocumentType(documentType);
        var baseName = GetAutosaveBaseName();

        return new AutosaveContext(
            baseName,
            documentType,
            extension,
            _currentFilePath,
            string.IsNullOrEmpty(_currentFilePath));
    }

    private DocumentType GetCurrentDocumentTypeForSaving()
    {
        if (!string.IsNullOrEmpty(_currentFilePath))
        {
            return DetermineDocumentType(_currentFilePath);
        }

        return _currentDocumentType;
    }

    private string GetAutosaveBaseName()
    {
        if (!string.IsNullOrEmpty(_currentFilePath))
        {
            return CreateDeterministicAutosaveName(_currentFilePath);
        }

        if (_untitledDocumentId == Guid.Empty)
        {
            _untitledDocumentId = Guid.NewGuid();
        }

        return _untitledDocumentId.ToString();
    }

    private static string CreateDeterministicAutosaveName(string path)
    {
        var sanitized = SanitizeFileName(Path.GetFileName(path));
        if (string.IsNullOrWhiteSpace(sanitized))
        {
            sanitized = "document";
        }

        using var sha = SHA256.Create();
        var hash = Convert.ToHexString(sha.ComputeHash(Encoding.UTF8.GetBytes(path))).Substring(0, 8);
        return $"{sanitized}-{hash}";
    }

    private static string SanitizeFileName(string name)
    {
        var invalidChars = Path.GetInvalidFileNameChars();
        var builder = new StringBuilder();

        foreach (char c in name)
        {
            builder.Append(invalidChars.Contains(c) ? '_' : c);
        }

        return builder.ToString();
    }

    private static void OnIsSpellCheckEnabledChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
    {
        if (d is MainWindow window)
        {
            window.OnSpellCheckEnabledChanged();
        }
    }

    private void OnSpellCheckEnabledChanged()
    {
        if (!_isLoadingSettings)
        {
            SaveUserSettings();
        }

        ConfigureCustomDictionary();
    }

    private void ConfigureCustomDictionary()
    {
        if (EditorBox == null)
        {
            return;
        }

        if (!IsSpellCheckEnabled || !EditorBox.SpellCheck.IsEnabled)
        {
            EditorBox.SpellCheck.CustomDictionaries.Clear();
            return;
        }

        EnsureCustomDictionaryFileExists();
        ReloadCustomDictionary();
    }

    private void EnsureCustomDictionaryFileExists()
    {
        try
        {
            var directory = Path.GetDirectoryName(_customDictionaryPath);
            if (!string.IsNullOrWhiteSpace(directory))
            {
                Directory.CreateDirectory(directory);
            }

            if (!File.Exists(_customDictionaryPath))
            {
                File.WriteAllText(_customDictionaryPath, string.Empty);
            }
        }
        catch
        {
            // Ignore custom dictionary failures to avoid interrupting the user.
        }
    }

    private void ReloadCustomDictionary()
    {
        if (EditorBox == null || !IsSpellCheckEnabled || !EditorBox.SpellCheck.IsEnabled)
        {
            return;
        }

        try
        {
            EditorBox.SpellCheck.CustomDictionaries.Clear();
            EditorBox.SpellCheck.CustomDictionaries.Add(new Uri(_customDictionaryPath, UriKind.Absolute));
        }
        catch
        {
            // Ignore custom dictionary failures to avoid interrupting the user.
        }
    }

    private void AddWordToCustomDictionary(string? word)
    {
        if (string.IsNullOrWhiteSpace(word))
        {
            return;
        }

        try
        {
            EnsureCustomDictionaryFileExists();
            File.AppendAllLines(_customDictionaryPath, new[] { word.Trim() });
        }
        catch
        {
            // Ignore custom dictionary failures to avoid interrupting the user.
        }
    }

    private void LoadUserSettings()
    {
        var settings = new UserSettings();

        try
        {
            if (File.Exists(_settingsFilePath))
            {
                settings = JsonSerializer.Deserialize<UserSettings>(File.ReadAllText(_settingsFilePath)) ?? settings;
            }
        }
        catch
        {
            // Ignore settings load failures to avoid blocking startup.
        }

        _isLoadingSettings = true;
        try
        {
            SetCurrentValue(IsSpellCheckEnabledProperty, settings.IsSpellCheckEnabled);
        }
        finally
        {
            _isLoadingSettings = false;
        }
    }

    private void SaveUserSettings()
    {
        try
        {
            var directory = Path.GetDirectoryName(_settingsFilePath);
            if (!string.IsNullOrWhiteSpace(directory))
            {
                Directory.CreateDirectory(directory);
            }

            var settings = new UserSettings
            {
                IsSpellCheckEnabled = IsSpellCheckEnabled
            };
            File.WriteAllText(_settingsFilePath, JsonSerializer.Serialize(settings, new JsonSerializerOptions
            {
                WriteIndented = true
            }));
        }
        catch
        {
            // Ignore settings save failures to avoid interrupting the user.
        }
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

    private void InsertTimestampMenuItem_Click(object sender, RoutedEventArgs e)
    {
        InsertTimestampAtCaret();
    }

    private void WordCountMenuItem_Click(object sender, RoutedEventArgs e)
    {
        string selectionText = EditorBox?.Selection.Text ?? string.Empty;
        string documentText = GetEditorPlainText();

        string normalizedSelection = NormalizeTextForCounting(selectionText);
        string normalizedDocument = NormalizeTextForCounting(documentText);

        var (selectionWordCount, selectionCharCount) = string.IsNullOrEmpty(selectionText)
            ? (0, 0)
            : CountWordsAndCharacters(normalizedSelection);

        var (documentWordCount, documentCharCount) = CountWordsAndCharacters(normalizedDocument);

        string message =
            $"Selection: {selectionWordCount} words, {selectionCharCount} characters\n" +
            $"Document: {documentWordCount} words, {documentCharCount} characters";

        System.Windows.MessageBox.Show(this, message, "Word count", MessageBoxButton.OK, MessageBoxImage.Information);
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

            System.Windows.MessageBox.Show(this,
                "PDF export completed:\n" + pdfPath,
                "Export as PDF",
                MessageBoxButton.OK,
                MessageBoxImage.None);
        }
        catch (Exception ex)
        {
            System.Windows.MessageBox.Show(this,
                "Failed to export PDF.\n\n" + ex.Message,
                "Export as PDF",
                MessageBoxButton.OK,
                MessageBoxImage.None);
        }
    }

    private void ExportMarkdownMenuItem_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            var dialog = new Microsoft.Win32.SaveFileDialog
            {
                Title = "Export as Markdown",
                Filter = "Markdown Files (*.md)|*.md|All Files (*.*)|*.*",
                DefaultExt = ".md",
                FileName = !string.IsNullOrEmpty(_currentFilePath)
                    ? Path.GetFileNameWithoutExtension(_currentFilePath) + ".md"
                    : "Document.md",
            };

            bool? result = dialog.ShowDialog(this);
            if (result != true)
                return;

            string markdownPath = dialog.FileName;

            if (!TryWriteDocument(markdownPath, DocumentType.Markdown, showErrors: true, applyMarkdownMode: false))
                return;

            System.Windows.MessageBox.Show(this,
                "Markdown export completed:\n" + markdownPath,
                "Export as Markdown",
                MessageBoxButton.OK,
                MessageBoxImage.None);
        }
        catch (Exception ex)
        {
            System.Windows.MessageBox.Show(this,
                "Failed to export Markdown.\n\n" + ex.Message,
                "Export as Markdown",
                MessageBoxButton.OK,
                MessageBoxImage.None);
        }
    }

    private void ExportOdtMenuItem_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            var dialog = new Microsoft.Win32.SaveFileDialog
            {
                Title = "Export as ODT",
                Filter = "OpenDocument Text (*.odt)|*.odt|All Files (*.*)|*.*",
                DefaultExt = ".odt",
                FileName = !string.IsNullOrEmpty(_currentFilePath)
                    ? Path.GetFileNameWithoutExtension(_currentFilePath) + ".odt"
                    : "Document.odt",
            };

            bool? result = dialog.ShowDialog(this);
            if (result != true)
                return;

            string odtPath = dialog.FileName;
            ExportDocumentToOdt(odtPath);

            System.Windows.MessageBox.Show(this,
                "ODT export completed:\n" + odtPath,
                "Export as ODT",
                MessageBoxButton.OK,
                MessageBoxImage.None);
        }
        catch (Exception ex)
        {
            System.Windows.MessageBox.Show(this,
                "Failed to export ODT.\n\n" + ex.Message,
                "Export as ODT",
                MessageBoxButton.OK,
                MessageBoxImage.None);
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
                bool isItalic = run.FontStyle == System.Windows.FontStyles.Italic;
                bool isUnderline = run.TextDecorations?.Contains(TextDecorations.Underline[0]) == true;
                bool isStrikethrough = run.TextDecorations?.Contains(TextDecorations.Strikethrough[0]) == true;
                bool isMonospaced = IsMonospacedFont(run.FontFamily);

                pdfParagraph.Runs.Add(new PdfTextRun
                {
                    Text = text,
                    IsBold = isBold,
                    IsItalic = isItalic,
                    IsUnderline = isUnderline,
                    IsStrikethrough = isStrikethrough,
                    IsMonospaced = isMonospaced
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

        double fontSize = 12;
        const double lineSpacingFactor = 1.4;
        var fontCache = new Dictionary<(string Family, XFontStyleEx Style), XFont>();
        XFont regularFont = GetOrCreateFont(
            fontCache,
            _styleConfiguration.BodyFontFamily,
            fontSize,
            XFontStyleEx.Regular);

        // Layout: 1-inch margins (72 points per inch)
        double marginLeft = 72;
        double marginTop = 72;
        double marginRight = 72;
        double marginBottom = 72;

        double y = marginTop;
        foreach (var paragraph in paragraphs)
        {
            double usableWidth = page.Width.Point - marginLeft - marginRight;

            var wrappedLines = WrapParagraphRuns(
                paragraph.Runs,
                gfx,
                usableWidth,
                run => GetFontForRun(run, fontCache, _styleConfiguration, fontSize));

            double lastLineHeight = 0;

            foreach (var line in wrappedLines)
            {
                double lineHeight = GetLineHeight(line, gfx, lineSpacingFactor, regularFont);

                if (y + lineHeight > page.Height.Point - marginBottom)
                {
                    // New page
                    page = document.AddPage();
                    gfx = XGraphics.FromPdfPage(page);
                    y = marginTop;
                }

                if (line.Count == 0)
                {
                    y += lineHeight;
                    lastLineHeight = lineHeight;
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
                lastLineHeight = lineHeight;
            }

            // Extra space between paragraphs
            double paragraphSpacing = (lastLineHeight > 0
                ? lastLineHeight
                : GetLineHeight(Array.Empty<PdfLineSpan>(), gfx, lineSpacingFactor, regularFont)) * 0.5;
            y += paragraphSpacing;
        }

        document.Save(pdfPath);
    }

    private double GetLineHeight(
        IReadOnlyCollection<PdfLineSpan> line,
        XGraphics gfx,
        double lineSpacingFactor,
        XFont fallbackFont)
    {
        double maxHeight = 0;

        foreach (var span in line)
        {
            double spanHeight = span.Font.GetHeight();
            maxHeight = Math.Max(maxHeight, spanHeight);
        }

        double baseHeight = maxHeight > 0
            ? maxHeight
            : fallbackFont.GetHeight();

        return baseHeight * lineSpacingFactor;
    }

    private List<List<PdfLineSpan>> WrapParagraphRuns(
        IEnumerable<PdfTextRun> runs,
        XGraphics gfx,
        double maxWidth,
        Func<PdfTextRun, XFont> fontSelector)
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
            XFont font = fontSelector(run);
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

    private static string GetFontFamilyForRun(PdfTextRun run, StyleConfiguration styleConfiguration)
    {
        return run.IsMonospaced
            ? styleConfiguration.MonoFontFamily
            : styleConfiguration.BodyFontFamily;
    }

    private static XFont GetFontForRun(
        PdfTextRun run,
        Dictionary<(string Family, XFontStyleEx Style), XFont> fontCache,
        StyleConfiguration styleConfiguration,
        double fontSize)
    {
        XFontStyleEx style = XFontStyleEx.Regular;

        if (run.IsBold)
            style |= XFontStyleEx.Bold;
        if (run.IsItalic)
            style |= XFontStyleEx.Italic;
        if (run.IsUnderline)
            style |= XFontStyleEx.Underline;
        if (run.IsStrikethrough)
            style |= XFontStyleEx.Strikeout;

        string selectedFamily = GetFontFamilyForRun(run, styleConfiguration);

        return GetOrCreateFont(fontCache, selectedFamily, fontSize, style);
    }

    private static XFont GetOrCreateFont(
        Dictionary<(string Family, XFontStyleEx Style), XFont> fontCache,
        string fontFamily,
        double fontSize,
        XFontStyleEx style)
    {
        var key = (fontFamily, style);

        if (!fontCache.TryGetValue(key, out var font))
        {
            font = new XFont(fontFamily, fontSize, style);
            fontCache[key] = font;
        }

        return font;
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

    private void ExportDocumentToOdt(string odtPath)
    {
        var paragraphs = ExtractParagraphsForPdf();

        XNamespace office = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
        XNamespace style = "urn:oasis:names:tc:opendocument:xmlns:style:1.0";
        XNamespace text = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";
        XNamespace fo = "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0";
        XNamespace svg = "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0";
        XNamespace manifest = "urn:oasis:names:tc:opendocument:xmlns:manifest:1.0";

        var automaticStyles = new List<XElement>();

        // When serializing to ODT we emit two named text styles: Text (body font) and Code (mono
        // font).  Runs marked IsMonospaced are mapped to Code; everything else uses Text.  The font
        // faces are written into styles.xml so consumers that lack the primary font can still fall
        // back to the Japanese-friendly alternates listed in StyleConfiguration.
        var bodyElements = BuildOdtBodyElements(paragraphs, automaticStyles, text, style, fo).ToList();

        var contentDoc = new XDocument(
            new XDeclaration("1.0", "UTF-8", "yes"),
            new XElement(office + "document-content",
                new XAttribute(XNamespace.Xmlns + "office", office),
                new XAttribute(XNamespace.Xmlns + "style", style),
                new XAttribute(XNamespace.Xmlns + "text", text),
                new XAttribute(XNamespace.Xmlns + "fo", fo),
                new XElement(office + "automatic-styles", automaticStyles),
                new XElement(office + "body",
                    new XElement(office + "text", bodyElements))));

        var stylesDoc = new XDocument(
            new XDeclaration("1.0", "UTF-8", "yes"),
            new XElement(office + "document-styles",
                new XAttribute(XNamespace.Xmlns + "office", office),
                new XAttribute(XNamespace.Xmlns + "style", style),
                new XAttribute(XNamespace.Xmlns + "svg", svg),
                new XElement(office + "font-face-decls",
                    new XElement(style + "font-face",
                        new XAttribute(style + "name", "BodyFont"),
                        new XAttribute(svg + "font-family", _styleConfiguration.BodyFontFamily)),
                    new XElement(style + "font-face",
                        new XAttribute(style + "name", "MonoFont"),
                        new XAttribute(svg + "font-family", _styleConfiguration.MonoFontFamily))),
                new XElement(office + "styles",
                    new XElement(style + "style",
                        new XAttribute(style + "name", "Text"),
                        new XAttribute(style + "family", "text"),
                        new XElement(style + "text-properties",
                            new XAttribute(style + "font-name", "BodyFont"))),
                    new XElement(style + "style",
                        new XAttribute(style + "name", "Code"),
                        new XAttribute(style + "family", "text"),
                        new XElement(style + "text-properties",
                            new XAttribute(style + "font-name", "MonoFont"))))));

        var manifestDoc = new XDocument(
            new XDeclaration("1.0", "UTF-8", "yes"),
            new XElement(manifest + "manifest",
                new XAttribute(XNamespace.Xmlns + "manifest", manifest),
                new XAttribute(manifest + "version", "1.2"),
                new XElement(manifest + "file-entry",
                    new XAttribute(manifest + "full-path", "/"),
                    new XAttribute(manifest + "media-type", "application/vnd.oasis.opendocument.text")),
                new XElement(manifest + "file-entry",
                    new XAttribute(manifest + "full-path", "content.xml"),
                    new XAttribute(manifest + "media-type", "text/xml")),
                new XElement(manifest + "file-entry",
                    new XAttribute(manifest + "full-path", "styles.xml"),
                    new XAttribute(manifest + "media-type", "text/xml"))));

        using var archive = ZipFile.Open(odtPath, ZipArchiveMode.Create);

        var mimeEntry = archive.CreateEntry("mimetype", CompressionLevel.NoCompression);
        using (var mimeWriter = new StreamWriter(mimeEntry.Open(), new UTF8Encoding(false)))
        {
            mimeWriter.Write("application/vnd.oasis.opendocument.text");
        }

        WriteXmlEntry(archive, "content.xml", contentDoc);
        WriteXmlEntry(archive, "styles.xml", stylesDoc);
        WriteXmlEntry(archive, "META-INF/manifest.xml", manifestDoc);
    }

    private static void WriteXmlEntry(ZipArchive archive, string entryName, XDocument document)
    {
        var entry = archive.CreateEntry(entryName, CompressionLevel.Optimal);
        using var entryStream = entry.Open();
        using var writer = new StreamWriter(entryStream, new UTF8Encoding(false));
        document.Save(writer);
    }

    private IEnumerable<XElement> BuildOdtBodyElements(
        IReadOnlyList<PdfParagraph> paragraphs,
        List<XElement> automaticStyles,
        XNamespace text,
        XNamespace style,
        XNamespace fo)
    {
        // We parse each paragraph to detect leading list markers (numeric, alphabetic, or bullet)
        // and their indentation level.  When a marker is present we emit nested text:list elements
        // so consumers see true list structure; otherwise we fall back to plain text:p elements
        // with the original prefix intact.  This mapping keeps manual list formatting reversible:
        // office suites that ignore text:list still receive the exact marker/spacing, while suites
        // that honor it can recover proper numbering and nesting.
        var result = new List<XElement>();
        var listStack = new List<OdtListState>();
        var listStyles = new Dictionary<string, XElement>();

        foreach (PdfParagraph paragraph in paragraphs)
        {
            if (TryParseListParagraph(paragraph, out ParsedListLine parsedLine, out List<PdfTextRun> contentRuns))
            {
                AdjustListStackForLine(parsedLine, result, listStack, listStyles, automaticStyles, text, style);

                OdtListState currentLevel = listStack[parsedLine.IndentLevel];
                var listItem = new XElement(text + "list-item");
                listItem.Add(CreateOdtParagraph(contentRuns, automaticStyles, text, style, fo));

                currentLevel.ListElement.Add(listItem);
                currentLevel.LastItem = listItem;
            }
            else
            {
                listStack.Clear();
                result.Add(CreateOdtParagraph(paragraph.Runs, automaticStyles, text, style, fo));
            }
        }

        return result;
    }

    private void AdjustListStackForLine(
        ParsedListLine parsedLine,
        List<XElement> rootElements,
        List<OdtListState> stack,
        Dictionary<string, XElement> styleCache,
        List<XElement> automaticStyles,
        XNamespace text,
        XNamespace style)
    {
        int targetDepth = parsedLine.IndentLevel + 1;

        while (stack.Count > targetDepth)
            stack.RemoveAt(stack.Count - 1);

        string styleName = GetOrCreateListStyleName(parsedLine, stack.Count, styleCache, automaticStyles, text, style);

        if (stack.Count == 0 || !stack[^1].Matches(parsedLine, styleName))
        {
            stack.Add(CreateListLevel(parsedLine, styleName, stack.LastOrDefault(), rootElements, text));
        }

        while (stack.Count < targetDepth)
        {
            stack.Add(CreateListLevel(parsedLine, styleName, stack.LastOrDefault(), rootElements, text));
        }

        if (!stack[^1].Matches(parsedLine, styleName))
        {
            stack[^1] = CreateListLevel(parsedLine, styleName, stack.Count > 1 ? stack[^2] : null, rootElements, text);
        }
    }

    private OdtListState CreateListLevel(
        ParsedListLine parsedLine,
        string styleName,
        OdtListState? parent,
        List<XElement> rootElements,
        XNamespace text)
    {
        var listElement = new XElement(text + "list",
            new XAttribute(text + "style-name", styleName));

        if (parent?.LastItem != null)
            parent.LastItem.Add(listElement);
        else
            rootElements.Add(listElement);

        return new OdtListState
        {
            Type = parsedLine.Type,
            IsUppercaseLetter = parsedLine.IsUppercaseLetter,
            BulletSymbol = parsedLine.Marker,
            ListElement = listElement,
            StyleName = styleName
        };
    }

    private string GetOrCreateListStyleName(
        ParsedListLine parsedLine,
        int currentDepth,
        Dictionary<string, XElement> styleCache,
        List<XElement> automaticStyles,
        XNamespace text,
        XNamespace style)
    {
        string key = $"{parsedLine.Type}:{parsedLine.IsUppercaseLetter}:{parsedLine.Marker}:{parsedLine.Punctuation}:{parsedLine.Spacing}";

        if (!styleCache.TryGetValue(key, out XElement? styleElement))
        {
            string styleName = $"GeneratedListStyle{styleCache.Count + 1}";
            styleElement = new XElement(text + "list-style",
                new XAttribute(style + "name", styleName));
            styleCache[key] = styleElement;
            automaticStyles.Add(styleElement);
        }

        int level = currentDepth + 1;
        bool hasLevel = styleElement.Elements()
            .Any(e => int.TryParse((string?)e.Attribute(text + "level"), out int parsedLevel) && parsedLevel == level);

        if (!hasLevel)
        {
            string suffix = parsedLine.Punctuation + parsedLine.Spacing;
            XElement levelElement = parsedLine.Type switch
            {
                ListLineType.Numbered => new XElement(text + "list-level-style-number",
                    new XAttribute(text + "level", level),
                    new XAttribute(style + "num-format", "1"),
                    new XAttribute(style + "num-suffix", suffix)),
                ListLineType.Lettered => new XElement(text + "list-level-style-number",
                    new XAttribute(text + "level", level),
                    new XAttribute(style + "num-format", parsedLine.IsUppercaseLetter ? "A" : "a"),
                    new XAttribute(style + "num-suffix", suffix)),
                _ => new XElement(text + "list-level-style-bullet",
                    new XAttribute(text + "level", level),
                    new XAttribute(text + "bullet-char", parsedLine.Marker),
                    new XAttribute(style + "num-suffix", suffix))
            };

            styleElement.Add(levelElement);
        }

        return (string)styleElement.Attribute(style + "name")!;
    }

    private bool TryParseListParagraph(
        PdfParagraph paragraph,
        out ParsedListLine parsedLine,
        out List<PdfTextRun> contentRuns)
    {
        string paragraphText = string.Concat(paragraph.Runs.Select(r => r.Text));

        if (!ListParser.TryParseListLine(paragraphText, out parsedLine))
        {
            contentRuns = paragraph.Runs.ToList();
            return false;
        }

        int prefixLength = parsedLine.IndentText.Length
            + parsedLine.Marker.Length
            + parsedLine.Punctuation.Length
            + parsedLine.Spacing.Length;

        contentRuns = TrimPrefixFromRuns(paragraph.Runs, prefixLength);
        return true;
    }

    private List<PdfTextRun> TrimPrefixFromRuns(IReadOnlyList<PdfTextRun> runs, int prefixLength)
    {
        var result = new List<PdfTextRun>();
        int remaining = prefixLength;
        bool started = remaining <= 0;

        foreach (PdfTextRun run in runs)
        {
            string text = run.Text;

            if (!started)
            {
                if (text.Length <= remaining)
                {
                    remaining -= text.Length;
                    continue;
                }

                string trimmedText = text[remaining..];
                result.Add(CloneRunWithText(run, trimmedText));
                started = true;
            }
            else
            {
                result.Add(CloneRunWithText(run, text));
            }
        }

        return result;
    }

    private static PdfTextRun CloneRunWithText(PdfTextRun source, string text)
    {
        return new PdfTextRun
        {
            Text = text,
            IsBold = source.IsBold,
            IsItalic = source.IsItalic,
            IsUnderline = source.IsUnderline,
            IsStrikethrough = source.IsStrikethrough,
            IsMonospaced = source.IsMonospaced
        };
    }

    private sealed class OdtListState
    {
        public ListLineType Type { get; init; }
        public bool IsUppercaseLetter { get; init; }
        public string BulletSymbol { get; init; } = string.Empty;
        public XElement ListElement { get; init; } = null!;
        public XElement? LastItem { get; set; }
        public string StyleName { get; init; } = string.Empty;

        public bool Matches(ParsedListLine parsedLine, string styleName)
        {
            return Type == parsedLine.Type
                && IsUppercaseLetter == parsedLine.IsUppercaseLetter
                && BulletSymbol == parsedLine.Marker
                && string.Equals(StyleName, styleName, StringComparison.Ordinal);
        }
    }

    private XElement CreateOdtParagraph(
        IReadOnlyList<PdfTextRun> runs,
        List<XElement> automaticStyles,
        XNamespace text,
        XNamespace style,
        XNamespace fo)
    {
        var paragraphElement = new XElement(text + "p");

        foreach (var run in runs)
        {
            AppendSpanElements(paragraphElement, run, automaticStyles, text, style, fo);
        }

        return paragraphElement;
    }

    private void AppendSpanElements(
        XElement paragraphElement,
        PdfTextRun run,
        List<XElement> automaticStyles,
        XNamespace text,
        XNamespace style,
        XNamespace fo)
    {
        string styleName = GetSpanStyleName(run, automaticStyles, style, fo);

        string normalized = run.Text.Replace("\r\n", "\n");
        string[] parts = normalized.Split('\n');

        for (int i = 0; i < parts.Length; i++)
        {
            if (!string.IsNullOrEmpty(parts[i]))
            {
                paragraphElement.Add(new XElement(text + "span",
                    new XAttribute(text + "style-name", styleName),
                    parts[i]));
            }

            if (i < parts.Length - 1)
            {
                paragraphElement.Add(new XElement(text + "line-break"));
            }
        }
    }

    private string GetSpanStyleName(
        PdfTextRun run,
        List<XElement> automaticStyles,
        XNamespace style,
        XNamespace fo)
    {
        string baseStyle = run.IsMonospaced ? "Code" : "Text";

        if (!run.IsBold && !run.IsItalic && !run.IsUnderline)
            return baseStyle;

        string styleName = baseStyle
            + (run.IsBold ? "Bold" : string.Empty)
            + (run.IsItalic ? "Italic" : string.Empty)
            + (run.IsUnderline ? "Underline" : string.Empty);

        bool exists = automaticStyles.Any(s => (string?)s.Attribute(style + "name") == styleName);
        if (exists)
            return styleName;

        var properties = new List<XAttribute>();

        if (run.IsBold)
            properties.Add(new XAttribute(fo + "font-weight", "bold"));
        if (run.IsItalic)
            properties.Add(new XAttribute(fo + "font-style", "italic"));
        if (run.IsUnderline)
        {
            properties.Add(new XAttribute(style + "text-underline-style", "solid"));
            properties.Add(new XAttribute(style + "text-underline-type", "single"));
            properties.Add(new XAttribute(style + "text-underline-width", "auto"));
            properties.Add(new XAttribute(style + "text-underline-color", "font-color"));
        }

        automaticStyles.Add(new XElement(style + "style",
            new XAttribute(style + "name", styleName),
            new XAttribute(style + "family", "text"),
            new XAttribute(style + "parent-style-name", baseStyle),
            new XElement(style + "text-properties", properties)));

        return styleName;
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

        _autosaveService.DeleteCurrentAutosaveFiles();
        _isLoadingDocument = true;
        try
        {
            EditorBox.Document = CreateStyledDocument();
        }
        finally
        {
            _isLoadingDocument = false;
        }

        var newDocumentType = _isMarkdownMode ? DocumentType.Markdown : DocumentType.RichText;
        SetCurrentFile(null, newDocumentType);
        MarkClean();
        UpdateFormattingControls();
        ApplyTheme();
        ResetZoomStateForDocument();
        ApplyZoom();
        UpdateMarkdownPreview();
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

        _autosaveService.DeleteCurrentAutosaveFiles();

        if (!File.Exists(path))
        {
            System.Windows.MessageBox.Show(this,
                $"The file could not be found:\n{path}",
                "File Not Found",
                MessageBoxButton.OK,
                MessageBoxImage.None);
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

        var dialog = new Microsoft.Win32.OpenFileDialog
        {
            Filter = "Text Documents (*.txt)|*.txt|Rich Text Format (*.rtf)|*.rtf|Markdown Files (*.md)|*.md|OpenDocument Text (*.odt)|*.odt|All files (*.*)|*.*"
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
            DocumentType.OpenDocument => ".odt",
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
            DocumentType.OpenDocument => 4,
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
                var text = File.ReadAllText(path);
                SetEditorPlainText(text);
                SetMarkdownMode(true);
            }
            else if (documentType == DocumentType.OpenDocument)
            {
                LoadOdtIntoEditor(path);
                SetMarkdownMode(false);
            }
            else if (documentType == DocumentType.RichText)
            {
                LoadRtfIntoEditor(path);
                SetMarkdownMode(false);
            }
            else
            {
                var text = File.ReadAllText(path);
                SetEditorPlainText(text);
                DisableMarkdownModeLayout();
            }

            SetCurrentFile(path, documentType);
            UpdateFormattingControls();
            MarkClean();
        }
        catch (Exception ex)
        {
            System.Windows.MessageBox.Show($"Failed to load document: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.None);
        }
        finally
        {
            _isLoadingDocument = false;
        }
    }

    public void LoadRecoveredDocument(RecoveryItem recoveryItem)
    {
        if (!File.Exists(recoveryItem.AutosavePath))
        {
            System.Windows.MessageBox.Show(
                $"The autosave file could not be found:\n{recoveryItem.AutosavePath}",
                "Autosave Missing",
                MessageBoxButton.OK,
                MessageBoxImage.None);
            return;
        }

        try
        {
            _isLoadingDocument = true;
            _isDirty = false;
            var documentType = recoveryItem.DocumentType;

            if (documentType == DocumentType.Markdown)
            {
                var text = File.ReadAllText(recoveryItem.AutosavePath);
                SetEditorPlainText(text);
                SetMarkdownMode(true);
            }
            else if (documentType == DocumentType.OpenDocument)
            {
                LoadOdtIntoEditor(recoveryItem.AutosavePath);
                SetMarkdownMode(false);
            }
            else if (documentType == DocumentType.RichText)
            {
                LoadRtfIntoEditor(recoveryItem.AutosavePath);
                SetMarkdownMode(false);
            }
            else
            {
                var text = File.ReadAllText(recoveryItem.AutosavePath);
                SetEditorPlainText(text);
                DisableMarkdownModeLayout();
            }

            _currentFilePath = recoveryItem.IsUntitled ? null : recoveryItem.OriginalPath;
            _currentDocumentType = documentType;
            _previousDocumentType = documentType;
            _isRecoveredDocument = true;
            _pendingRecoveryAutosavePaths.Clear();
            _pendingRecoveryAutosavePaths.Add(recoveryItem.AutosavePath);
            _pendingRecoveryAutosavePaths.Add(recoveryItem.MetadataPath);

            if (recoveryItem.IsUntitled)
            {
                _untitledDocumentId = Guid.NewGuid();
            }
            else if (!string.IsNullOrWhiteSpace(recoveryItem.OriginalPath))
            {
                AddRecentDocument(recoveryItem.OriginalPath);
            }

            UpdateWindowTitle();
            UpdateFormattingControls();
            MarkDirty();
        }
        catch (Exception ex)
        {
            System.Windows.MessageBox.Show(
                $"Failed to recover the document: {ex.Message}",
                "Recovery Error",
                MessageBoxButton.OK,
                MessageBoxImage.None);
        }
        finally
        {
            _isLoadingDocument = false;
        }
    }

    private bool TryWriteDocument(string path, DocumentType documentType, bool showErrors, bool applyMarkdownMode)
    {
        try
        {
            if (documentType == DocumentType.RichText)
            {
                if (EditorBox != null)
                {
                    var range = new TextRange(EditorBox.Document.ContentStart, EditorBox.Document.ContentEnd);
                    using var stream = new FileStream(path, FileMode.Create, FileAccess.Write);
                    range.Save(stream, System.Windows.DataFormats.Rtf);
                }

                if (applyMarkdownMode)
                {
                    SetMarkdownMode(false);
                }
            }
            else if (documentType == DocumentType.Markdown)
            {
                var text = GetEditorPlainText();
                File.WriteAllText(path, text);

                if (applyMarkdownMode)
                {
                    SetMarkdownMode(true);
                }
            }
            else if (documentType == DocumentType.OpenDocument)
            {
                ExportDocumentToOdt(path);

                if (applyMarkdownMode)
                {
                    SetMarkdownMode(false);
                }
            }
            else
            {
                var text = GetEditorPlainText();
                File.WriteAllText(path, text);

                if (applyMarkdownMode)
                {
                    SetMarkdownMode(false);
                }
            }

            return true;
        }
        catch (Exception ex)
        {
            if (showErrors)
            {
                System.Windows.MessageBox.Show($"Failed to save document: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.None);
            }
            else
            {
                Debug.WriteLine($"Silent save failed for '{path}': {ex}");
            }

            return false;
        }
    }

    private bool SaveDocumentToFile(string path)
    {
        var documentType = DetermineDocumentType(path);
        if (!TryWriteDocument(path, documentType, showErrors: true, applyMarkdownMode: true))
        {
            return false;
        }

        SetCurrentFile(path, documentType);
        UpdateFormattingControls();
        MarkClean();
        _autosaveService.OnManualSave();
        return true;
    }

    private bool SaveAutosaveDocument(string path, DocumentType documentType)
    {
        return TryWriteDocument(path, documentType, showErrors: false, applyMarkdownMode: false);
    }

    private static DocumentType DetermineDocumentType(string path)
    {
        var extension = Path.GetExtension(path)?.ToLowerInvariant();
        return extension switch
        {
            ".rtf" => DocumentType.RichText,
            ".md" => DocumentType.Markdown,
            ".odt" => DocumentType.OpenDocument,
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
        var dialogDocumentType = _currentDocumentType == DocumentType.Markdown
            ? DocumentType.Markdown
            : (_isMarkdownMode ? DocumentType.Markdown : _currentDocumentType);

        var dialog = new Microsoft.Win32.SaveFileDialog
        {
            Filter = "Text Documents (*.txt)|*.txt|Rich Text Format (*.rtf)|*.rtf|Markdown Files (*.md)|*.md|OpenDocument Text (*.odt)|*.odt|All files (*.*)|*.*",
            AddExtension = true,
            DefaultExt = GetExtensionForDocumentType(dialogDocumentType),
            FilterIndex = GetFilterIndexForDocumentType(dialogDocumentType)
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

        var result = System.Windows.MessageBox.Show(
            "You have unsaved changes. Do you want to save them before continuing?",
            "Unsaved Changes",
            MessageBoxButton.YesNoCancel,
            MessageBoxImage.None);

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

    private void EditorBox_PreviewKeyDown(object sender, KeyEventArgs e)
    {
        if (EditorBox == null)
            return;

        bool isShiftTab = (Keyboard.Modifiers & ModifierKeys.Shift) == ModifierKeys.Shift;

        if (e.Key == Key.Enter)
        {
            if (HandleEnterInList())
            {
                e.Handled = true;
                return;
            }

            if (TryConvertCurrentParagraphToList(requireContentAfterPrefix: true, createFollowingListItem: true))
                e.Handled = true;
            return;
        }

        if (e.Key == Key.Tab)
        {
            if (HandleTabInList(isShiftTab))
                e.Handled = true;
            return;
        }

        if (e.Key == Key.Back)
        {
            if (HandleBackspaceAtListStart())
                e.Handled = true;
        }
    }

    private void EditorBox_ContextMenuOpening(object sender, ContextMenuEventArgs e)
    {
        if (EditorBox?.ContextMenu == null)
        {
            return;
        }

        var contextMenu = EditorBox.ContextMenu;

        for (int i = contextMenu.Items.Count - 1; i >= 0; i--)
        {
            if (contextMenu.Items[i] is FrameworkElement element
                && ReferenceEquals(element.Tag, SpellCheckContextMenuTag))
            {
                contextMenu.Items.RemoveAt(i);
            }
        }

        if (!IsSpellCheckEnabled || !EditorBox.SpellCheck.IsEnabled)
        {
            return;
        }

        var pointer = EditorBox.GetPositionFromPoint(Mouse.GetPosition(EditorBox), true)
            ?? EditorBox.CaretPosition;
        var spellingError = pointer == null
            ? null
            : EditorBox.GetSpellingError(pointer);

        if (spellingError == null)
        {
            return;
        }

        var suggestions = spellingError.Suggestions?.Cast<string>().Take(8).ToList()
            ?? new List<string>();
        var insertItems = new List<FrameworkElement>(suggestions.Count + 3);

        foreach (var suggestion in suggestions)
        {
            var suggestionItem = new MenuItem
            {
                Header = suggestion,
                Command = EditingCommands.CorrectSpellingError,
                CommandTarget = EditorBox,
                CommandParameter = suggestion,
                Tag = SpellCheckContextMenuTag
            };
            insertItems.Add(suggestionItem);
        }

        var ignoreItem = new MenuItem
        {
            Header = "Ignore",
            Command = EditingCommands.IgnoreSpellingError,
            CommandTarget = EditorBox,
            Tag = SpellCheckContextMenuTag
        };
        insertItems.Add(ignoreItem);

        var addToDictionaryItem = new MenuItem
        {
            Header = "Add to Dictionary",
            Tag = SpellCheckContextMenuTag,
            IsEnabled = false
        };
        insertItems.Add(addToDictionaryItem);

        insertItems.Add(new Separator { Tag = SpellCheckContextMenuTag });

        var insertIndex = 0;
        foreach (var item in insertItems)
        {
            contextMenu.Items.Insert(insertIndex, item);
            insertIndex++;
        }
    }

    private bool GetCurrentLineTextAndOffsets(
        out string lineText,
        out TextPointer lineStart,
        out TextPointer lineEnd)
    {
        lineText = string.Empty;
        lineStart = lineEnd = EditorBox?.Document.ContentStart ?? new FlowDocument().ContentStart;

        if (EditorBox == null)
            return false;

        lineStart = EditorBox.CaretPosition.GetLineStartPosition(0)
            ?? EditorBox.Document.ContentStart;

        TextPointer? nextLineStart = lineStart.GetLineStartPosition(1);
        lineEnd = nextLineStart ?? EditorBox.Document.ContentEnd;
        lineText = new TextRange(lineStart, lineEnd).Text;

        return true;
    }

    private void ReplaceLineText(TextPointer lineStart, TextPointer lineEnd, string text)
    {
        new TextRange(lineStart, lineEnd).Text = text;
    }

    private sealed record SelectionLine(TextPointer Start, TextPointer End, string Text)
    {
        public string Content => Text.TrimEnd('\r', '\n');

        public string LineBreak => Text[Content.Length..];
    }

    private IEnumerable<SelectionLine> GetSelectedLines()
    {
        if (EditorBox?.Document == null)
            yield break;

        TextPointer selectionStart = EditorBox.Selection.Start;
        TextPointer selectionEnd = EditorBox.Selection.End;

        TextPointer currentLineStart = selectionStart.GetLineStartPosition(0) ?? selectionStart;
        TextPointer finalLineStart = selectionEnd.GetLineStartPosition(0) ?? selectionEnd;

        TextPointer? lineStart = currentLineStart;
        while (lineStart != null)
        {
            TextPointer? nextLineStart = lineStart.GetLineStartPosition(1);
            TextPointer lineEnd = nextLineStart ?? EditorBox.Document.ContentEnd;

            yield return new SelectionLine(lineStart, lineEnd, new TextRange(lineStart, lineEnd).Text);

            if (lineStart.CompareTo(finalLineStart) == 0)
                yield break;

            if (nextLineStart == null || nextLineStart.CompareTo(finalLineStart) > 0)
                lineStart = finalLineStart;
            else
                lineStart = nextLineStart;
        }
    }

    private static string GetLineIndentation(string text)
    {
        int index = 0;
        while (index < text.Length && char.IsWhiteSpace(text[index]))
            index++;

        return text[..index];
    }

    private Paragraph? GetCurrentParagraph(TextPointer caret)
    {
        return caret.Paragraph;
    }

    private ListItem? GetParentListItem(TextPointer caret)
    {
        TextElement? element = caret.Parent as TextElement;

        while (element != null)
        {
            if (element is ListItem listItem)
                return listItem;

            element = element.Parent as TextElement;
        }

        return null;
    }

    private static BlockCollection? GetParentBlockCollection(DependencyObject? element)
    {
        return element switch
        {
            FlowDocument document => document.Blocks,
            Section section => section.Blocks,
            ListItem listItem => listItem.Blocks,
            _ => null
        };
    }

    private bool TryConvertCurrentParagraphToList(bool requireContentAfterPrefix = false, bool createFollowingListItem = false)
    {
        if (EditorBox == null)
            return false;

        TextPointer caret = EditorBox.CaretPosition;
        Paragraph? paragraph = GetCurrentParagraph(caret);

        if (paragraph == null)
            return false;

        string paragraphText = new TextRange(paragraph.ContentStart, paragraph.ContentEnd).Text;
        string normalizedText = ListParser.NormalizeLine(paragraphText);

        if (!ListParser.TryParseListLine(normalizedText, out ParsedListLine parsedLine))
            return false;

        if (requireContentAfterPrefix && string.IsNullOrWhiteSpace(parsedLine.Content))
            return false;

        TextMarkerStyle markerStyle = parsedLine.Type switch
        {
            ListLineType.Numbered => TextMarkerStyle.Decimal,
            ListLineType.Lettered => parsedLine.IsUppercaseLetter ? TextMarkerStyle.UpperLatin : TextMarkerStyle.LowerLatin,
            _ => TextMarkerStyle.Disc
        };

        string prefixText = ListParser.BuildPrefix(
            parsedLine.IndentText,
            parsedLine.Marker,
            parsedLine.Punctuation,
            parsedLine.Spacing);

        if (!normalizedText.StartsWith(prefixText, StringComparison.Ordinal))
            return false;

        TextPointer? markerEnd = GetTextPointerAtTextOffset(paragraph.ContentStart, prefixText.Length);

        if (markerEnd == null)
            return false;

        int markerOffset = paragraph.ContentStart.GetOffsetToPosition(markerEnd);
        int caretOffset = paragraph.ContentStart.GetOffsetToPosition(caret);

        RemoveTextInRange(paragraph.ContentStart, markerEnd);

        List list = CreateList(markerStyle);

        BlockCollection? parentBlocks = GetParentBlockCollection(paragraph.Parent ?? (DependencyObject?)EditorBox?.Document);
        if (parentBlocks == null)
            parentBlocks = (paragraph.Parent as FlowDocument)?.Blocks;

        if (parentBlocks == null)
            return false;

        parentBlocks.InsertBefore(paragraph, list);
        parentBlocks.Remove(paragraph);
        ListItem listItem = new(paragraph);
        list.ListItems.Add(listItem);

        if (createFollowingListItem)
        {
            ListItem newItem = new(new Paragraph());
            list.ListItems.Add(newItem);
            EditorBox.CaretPosition = newItem.ContentStart;
        }
        else
        {
            int adjustedOffset = caretOffset >= markerOffset
                ? caretOffset - markerOffset
                : 0;
            TextPointer? adjustedCaret = GetTextPointerAtTextOffset(paragraph.ContentStart, adjustedOffset);
            EditorBox.CaretPosition = adjustedCaret ?? listItem.ContentStart;
        }
        MarkDirty();
        return true;
    }

    private static TextPointer? GetTextPointerAtTextOffset(TextPointer start, int offset)
    {
        if (offset <= 0)
            return start;

        TextPointer? current = start;
        int remaining = offset;

        while (current != null && remaining > 0)
        {
            TextPointerContext context = current.GetPointerContext(LogicalDirection.Forward);
            if (context == TextPointerContext.Text)
            {
                string text = current.GetTextInRun(LogicalDirection.Forward);
                int advance = Math.Min(text.Length, remaining);
                current = current.GetPositionAtOffset(advance, LogicalDirection.Forward);
                remaining -= advance;
            }
            else
            {
                current = current.GetNextContextPosition(LogicalDirection.Forward);
            }
        }

        return remaining == 0 ? current : null;
    }

    private static void RemoveTextInRange(TextPointer start, TextPointer end)
    {
        if (start.CompareTo(end) >= 0)
            return;

        TextPointer? current = start;

        while (current != null && current.CompareTo(end) < 0)
        {
            TextPointerContext context = current.GetPointerContext(LogicalDirection.Forward);
            if (context == TextPointerContext.Text)
            {
                Run? run = current.Parent as Run;
                if (run == null)
                {
                    current = current.GetNextContextPosition(LogicalDirection.Forward);
                    continue;
                }

                string runText = current.GetTextInRun(LogicalDirection.Forward);
                TextPointer? runEnd = current.GetPositionAtOffset(runText.Length, LogicalDirection.Forward);
                TextPointer segmentEnd = runEnd != null && runEnd.CompareTo(end) > 0 ? end : runEnd ?? end;

                if (segmentEnd.CompareTo(current) == 0)
                {
                    current = current.GetNextContextPosition(LogicalDirection.Forward);
                    continue;
                }

                int startOffset = run.ContentStart.GetOffsetToPosition(current);
                int endOffset = run.ContentStart.GetOffsetToPosition(segmentEnd);
                int removeLength = Math.Max(0, endOffset - startOffset);

                if (removeLength > 0 && startOffset >= 0 && startOffset < run.Text.Length)
                {
                    int safeLength = Math.Min(removeLength, run.Text.Length - startOffset);
                    if (safeLength > 0)
                        run.Text = run.Text.Remove(startOffset, safeLength);
                }

                current = segmentEnd;
            }
            else
            {
                current = current.GetNextContextPosition(LogicalDirection.Forward);
            }
        }
    }

    private bool HandleEnterInList()
    {
        if (EditorBox == null)
            return false;

        ListItem? currentItem = GetParentListItem(EditorBox.CaretPosition);
        if (currentItem == null)
            return false;

        if (IsListItemEmpty(currentItem))
        {
            return ExitListFromItem(currentItem);
        }

        Paragraph newParagraph = new();
        string trailingText = new TextRange(EditorBox.CaretPosition, currentItem.ContentEnd).Text;
        if (!string.IsNullOrEmpty(trailingText))
        {
            newParagraph.Inlines.Add(new Run(trailingText));
            new TextRange(EditorBox.CaretPosition, currentItem.ContentEnd).Text = string.Empty;
        }

        ListItem newItem = new(newParagraph);
        List parentList = (List)currentItem.Parent;
        parentList.ListItems.InsertAfter(currentItem, newItem);
        EditorBox.CaretPosition = newParagraph.ContentStart;
        MarkDirty();
        return true;
    }

    private static bool IsListItemEmpty(ListItem currentItem)
    {
        string itemText = new TextRange(currentItem.ContentStart, currentItem.ContentEnd).Text;
        if (!string.IsNullOrWhiteSpace(itemText))
            return false;

        List<Block> blocks = currentItem.Blocks.ToList();
        if (blocks.Count != 1 || blocks[0] is not Paragraph paragraph)
            return false;

        if (paragraph.Inlines.Count == 0)
            return true;

        return !InlineHasNonWhitespaceText(paragraph.Inlines);
    }

    private static bool InlineHasNonWhitespaceText(InlineCollection inlines)
    {
        foreach (Inline inline in inlines)
        {
            switch (inline)
            {
                case Run run:
                    if (!string.IsNullOrWhiteSpace(run.Text))
                        return true;
                    break;
                case Span span:
                    if (InlineHasNonWhitespaceText(span.Inlines))
                        return true;
                    break;
                default:
                    return true;
            }
        }

        return false;
    }

    private bool ExitListFromItem(ListItem currentItem)
    {
        List parentList = (List)currentItem.Parent;
        BlockCollection? outerBlocks = GetParentBlockCollection(parentList.Parent ?? (DependencyObject?)EditorBox?.Document);

        if (outerBlocks == null)
            outerBlocks = (parentList.Parent as FlowDocument)?.Blocks;

        if (outerBlocks == null)
            return false;

        List<ListItem> items = parentList.ListItems.Cast<ListItem>().ToList();
        int indexInList = items.IndexOf(currentItem);
        int totalItems = items.Count;

        Paragraph replacementParagraph = new();

        if (totalItems == 1)
        {
            outerBlocks.InsertAfter(parentList, replacementParagraph);
            outerBlocks.Remove(parentList);
        }
        else if (indexInList == 0)
        {
            parentList.ListItems.Remove(currentItem);
            outerBlocks.InsertBefore(parentList, replacementParagraph);
        }
        else if (indexInList == totalItems - 1)
        {
            parentList.ListItems.Remove(currentItem);
            outerBlocks.InsertAfter(parentList, replacementParagraph);
        }
        else
        {
            List tailList = new()
            {
                MarkerStyle = parentList.MarkerStyle
            };

            parentList.ListItems.Remove(currentItem);

            foreach (ListItem movedItem in items.Skip(indexInList + 1))
            {
                parentList.ListItems.Remove(movedItem);
                tailList.ListItems.Add(movedItem);
            }

            outerBlocks.InsertAfter(parentList, replacementParagraph);
            outerBlocks.InsertAfter(replacementParagraph, tailList);
        }

        EditorBox.CaretPosition = replacementParagraph.ContentStart;
        MarkDirty();
        return true;
    }

    private bool HandleTabInList(bool isShift)
    {
        if (EditorBox == null)
            return false;

        ListItem? currentItem = GetParentListItem(EditorBox.CaretPosition);
        if (currentItem == null)
            return false;

        List parentList = (List)currentItem.Parent;
        List<ListItem> items = parentList.ListItems.Cast<ListItem>().ToList();
        int index = items.IndexOf(currentItem);

        if (isShift)
        {
            ListItem? parentListItem = parentList.Parent as ListItem;
            List? outerList = parentListItem?.Parent as List;

            if (outerList == null || parentListItem == null)
                return false;

            parentList.ListItems.Remove(currentItem);
            outerList.ListItems.InsertAfter(parentListItem, currentItem);

            if (parentList.ListItems.Count == 0)
                parentListItem.Blocks.Remove(parentList);

            EditorBox.CaretPosition = currentItem.ContentStart;
            MarkDirty();
            return true;
        }

        if (index <= 0)
            return false;

        ListItem previousItem = items[index - 1];
        List? nestedList = previousItem.Blocks.OfType<List>().LastOrDefault();

        if (nestedList == null)
        {
            nestedList = new List
            {
                MarkerStyle = parentList.MarkerStyle
            };
            previousItem.Blocks.Add(nestedList);
        }

        parentList.ListItems.Remove(currentItem);
        nestedList.ListItems.Add(currentItem);
        EditorBox.CaretPosition = currentItem.ContentStart;
        MarkDirty();
        return true;
    }

    private bool HandleBackspaceAtListStart()
    {
        if (EditorBox == null)
            return false;

        ListItem? currentItem = GetParentListItem(EditorBox.CaretPosition);
        if (currentItem == null)
            return false;

        if (EditorBox.CaretPosition.CompareTo(currentItem.ContentStart) > 0)
            return false;

        return ConvertListItemToParagraph(currentItem);
    }

    private bool ConvertListItemToParagraph(ListItem currentItem)
    {
        List parentList = (List)currentItem.Parent;
        BlockCollection? outerBlocks = GetParentBlockCollection(parentList.Parent ?? (DependencyObject?)EditorBox?.Document);

        if (outerBlocks == null)
            outerBlocks = (parentList.Parent as FlowDocument)?.Blocks;

        if (outerBlocks == null)
            return false;

        string itemText = new TextRange(currentItem.ContentStart, currentItem.ContentEnd).Text;
        Paragraph paragraph = new(new Run(itemText));

        List<ListItem> items = parentList.ListItems.Cast<ListItem>().ToList();
        int indexInList = items.IndexOf(currentItem);
        int totalItems = items.Count;

        if (totalItems == 1)
        {
            outerBlocks.InsertAfter(parentList, paragraph);
            outerBlocks.Remove(parentList);
        }
        else if (indexInList == 0)
        {
            parentList.ListItems.Remove(currentItem);
            outerBlocks.InsertBefore(parentList, paragraph);
        }
        else if (indexInList == totalItems - 1)
        {
            parentList.ListItems.Remove(currentItem);
            outerBlocks.InsertAfter(parentList, paragraph);
        }
        else
        {
            List tailList = new()
            {
                MarkerStyle = parentList.MarkerStyle
            };

            parentList.ListItems.Remove(currentItem);

            foreach (ListItem movedItem in items.Skip(indexInList + 1))
            {
                parentList.ListItems.Remove(movedItem);
                tailList.ListItems.Add(movedItem);
            }

            outerBlocks.InsertAfter(parentList, paragraph);
            outerBlocks.InsertAfter(paragraph, tailList);
        }

        EditorBox.CaretPosition = paragraph.ContentStart;
        MarkDirty();
        return true;
    }

    private List<Paragraph> GetParagraphsFromSelection()
    {
        List<Paragraph> paragraphs = new();

        if (EditorBox == null)
            return paragraphs;

        TextPointer navigator = EditorBox.Selection.Start;
        TextPointer selectionEnd = EditorBox.Selection.End;

        while (navigator != null && navigator.CompareTo(selectionEnd) <= 0)
        {
            Paragraph? paragraph = navigator.Paragraph;
            if (paragraph != null && !paragraphs.Contains(paragraph))
                paragraphs.Add(paragraph);

            TextPointer? next = paragraph?.ContentEnd.GetNextInsertionPosition(LogicalDirection.Forward)
                ?? navigator.GetNextInsertionPosition(LogicalDirection.Forward);

            if (next == null || navigator.CompareTo(next) == 0)
                break;

            navigator = next;
        }

        return paragraphs;
    }

    private List<ListItem> GetListItemsFromSelection()
    {
        List<ListItem> items = new();

        if (EditorBox == null)
            return items;

        TextPointer navigator = EditorBox.Selection.Start;
        TextPointer selectionEnd = EditorBox.Selection.End;

        while (navigator != null && navigator.CompareTo(selectionEnd) <= 0)
        {
            ListItem? item = GetParentListItem(navigator);
            if (item != null && !items.Contains(item))
                items.Add(item);

            TextPointer? next = item?.ContentEnd.GetNextInsertionPosition(LogicalDirection.Forward)
                ?? navigator.GetNextInsertionPosition(LogicalDirection.Forward);

            if (next == null || navigator.CompareTo(next) == 0)
                break;

            navigator = next;
        }

        return items;
    }

    private void ApplyListStyle(TextMarkerStyle style)
    {
        if (EditorBox == null || !CanFormat())
            return;

        List<Paragraph> paragraphs = GetParagraphsFromSelection();
        if (paragraphs.Count == 0)
            return;

        bool allMatchingListItems = paragraphs.All(p =>
        {
            ListItem? item = GetParentListItem(p.ContentStart);
            return item != null && item.Parent is List list && list.MarkerStyle == style;
        });

        if (allMatchingListItems)
        {
            foreach (ListItem item in GetListItemsFromSelection())
                ConvertListItemToParagraph(item);
            return;
        }

        Paragraph firstParagraph = paragraphs[0];
        BlockCollection? parentBlocks = GetParentBlockCollection(firstParagraph.Parent ?? (DependencyObject?)EditorBox?.Document)
            ?? (firstParagraph.Parent as FlowDocument)?.Blocks;

        if (parentBlocks == null)
            return;

        List list = CreateList(style);

        parentBlocks.InsertBefore(firstParagraph, list);

        foreach (Paragraph paragraph in paragraphs)
        {
            parentBlocks.Remove(paragraph);
            list.ListItems.Add(new ListItem(paragraph));
        }

        EditorBox.CaretPosition = list.ListItems.Last().ContentStart;
        MarkDirty();
    }

    private static List CreateList(TextMarkerStyle style)
    {
        return new List
        {
            MarkerStyle = style,
            Margin = ListIndentationMargin,
            MarkerOffset = ListIndentationMarkerOffset,
            Padding = ListIndentationPadding
        };
    }

    private void RemoveListFormattingFromSelection()
    {
        foreach (ListItem item in GetListItemsFromSelection())
            ConvertListItemToParagraph(item);
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
        _autosaveService.Stop();
        _autosaveService.DeleteCurrentAutosaveFiles();
        DeletePendingRecoveryAutosaveFiles();
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
        UpdateFormattingControls();
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
                return;

            var selection = EditorBox.Selection;
            bool isSelectionEmpty = selection.IsEmpty;

            FontFamily? selectionFamily = GetSelectionOrCaretFontFamily(selection);
            if (selectionFamily != null)
            {
                FontFamily? matchedFamily = ResolveComboFontFamily(selectionFamily);
                FontFamilyComboBox.SelectedItem = matchedFamily;
                FontFamilyComboBox.Text = (matchedFamily ?? selectionFamily).Source;
                _lastKnownFontFamily = selectionFamily;
            }
            else
            {
                FontFamily? fallbackFamily = isSelectionEmpty
                    ? _lastKnownFontFamily ?? EditorBox.Document?.FontFamily
                    : null;

                if (fallbackFamily != null)
                {
                    _lastKnownFontFamily ??= fallbackFamily;
                    FontFamily? matchedFamily = ResolveComboFontFamily(fallbackFamily);
                    FontFamilyComboBox.SelectedItem = matchedFamily;
                    FontFamilyComboBox.Text = (matchedFamily ?? fallbackFamily).Source;
                }
                else
                {
                    FontFamilyComboBox.SelectedItem = null;
                    FontFamilyComboBox.Text = string.Empty;
                }
            }

            object sizeValue = selection.GetPropertyValue(Inline.FontSizeProperty);
            if (isSelectionEmpty && (sizeValue == DependencyProperty.UnsetValue || sizeValue is not double))
            {
                TextPointer? caret = EditorBox.CaretPosition;
                if (caret != null)
                {
                    sizeValue = new TextRange(caret, caret).GetPropertyValue(Inline.FontSizeProperty);
                }
            }

            double? selectionSize = sizeValue is double size ? DipsToPoints(size) : null;
            if (selectionSize != null)
            {
                double clampedSize = ClampFontSizeInPoints(selectionSize.Value);
                FontSizeComboBox.Text = clampedSize.ToString("0.#");
                var closest = _defaultFontSizes.FirstOrDefault(s => Math.Abs(s - clampedSize) < 0.1);
                FontSizeComboBox.SelectedItem = closest > 0 ? closest : null;
                _lastKnownFontSize = clampedSize;
            }
            else
            {
                double? fallbackSize = isSelectionEmpty
                    ? _lastKnownFontSize ?? (EditorBox.Document?.FontSize is double docSize ? DipsToPoints(docSize) : null)
                    : null;

                if (fallbackSize != null)
                {
                    double clampedSize = ClampFontSizeInPoints(fallbackSize.Value);
                    _lastKnownFontSize ??= clampedSize;
                    FontSizeComboBox.Text = clampedSize.ToString("0.#");
                    var closest = _defaultFontSizes.FirstOrDefault(s => Math.Abs(s - clampedSize) < 0.1);
                    FontSizeComboBox.SelectedItem = closest > 0 ? closest : null;
                }
                else
                {
                    FontSizeComboBox.SelectedItem = null;
                    FontSizeComboBox.Text = string.Empty;
                }
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

        if (ShouldUpdateDocumentFormattingDefault(Inline.FontFamilyProperty))
        {
            if (EditorBox.Document != null)
            {
                EditorBox.Document.FontFamily = fontFamily;
            }
        }
        else
        {
            ApplyPropertyToSelectionOrCaret(Inline.FontFamilyProperty, fontFamily);
        }

        _lastKnownFontFamily = fontFamily;
        UpdateFontControlsSelection(fontFamily, null);
    }

    private void ApplyFontSize(double fontSize)
    {
        if (EditorBox == null || !CanFormat())
            return;

        double clampedSize = ClampFontSizeInPoints(fontSize);
        if (ShouldUpdateDocumentFormattingDefault(Inline.FontSizeProperty))
        {
            if (EditorBox.Document != null)
            {
                EditorBox.Document.FontSize = PointsToDips(clampedSize);
            }
        }
        else
        {
            ApplyPropertyToSelectionOrCaret(Inline.FontSizeProperty, PointsToDips(clampedSize));
        }

        _lastKnownFontSize = clampedSize;
        UpdateFontControlsSelection(null, clampedSize);
    }

    private bool ShouldUpdateDocumentFormattingDefault(DependencyProperty property)
    {
        if (EditorBox == null)
            return false;

        TextSelection selection = EditorBox.Selection;
        if (!selection.IsEmpty)
            return false;

        object selectionValue = selection.GetPropertyValue(property);
        if (selectionValue != DependencyProperty.UnsetValue)
            return false;

        TextPointer? caret = EditorBox.CaretPosition;
        if (caret == null)
            return false;

        object caretValue = new TextRange(caret, caret).GetPropertyValue(property);
        return caretValue == DependencyProperty.UnsetValue;
    }

    private void ApplyPropertyToSelectionOrCaret(DependencyProperty property, object value)
    {
        if (EditorBox == null)
            return;

        TextSelection selection = EditorBox.Selection;

        EditorBox.BeginChange();
        try
        {
            if (selection.IsEmpty)
            {
                TextPointer? insertionPosition = EditorBox.CaretPosition?.GetInsertionPosition(LogicalDirection.Forward)
                    ?? EditorBox.CaretPosition;

                if (insertionPosition != null)
                {
                    selection.Select(insertionPosition, insertionPosition);
                    selection.ApplyPropertyValue(property, value);
                    EditorBox.CaretPosition = selection.End;
                }
            }
            else
            {
                selection.ApplyPropertyValue(property, value);
            }
        }
        finally
        {
            EditorBox.EndChange();
        }
    }

    private void UpdateFontControlsSelection(FontFamily? fontFamily, double? fontSize)
    {
        if (FontFamilyComboBox == null || FontSizeComboBox == null)
            return;

        _isUpdatingFontControls = true;
        try
        {
            if (fontFamily != null)
            {
                _lastKnownFontFamily = fontFamily;
                FontFamily? matchedFamily = ResolveComboFontFamily(fontFamily);
                FontFamilyComboBox.SelectedItem = matchedFamily;
                FontFamilyComboBox.Text = (matchedFamily ?? fontFamily).Source;
            }

            if (fontSize != null)
            {
                double clampedSize = ClampFontSizeInPoints(fontSize.Value);
                _lastKnownFontSize = clampedSize;
                FontSizeComboBox.Text = clampedSize.ToString("0.#");
                var closest = _defaultFontSizes.FirstOrDefault(s => Math.Abs(s - clampedSize) < 0.1);
                FontSizeComboBox.SelectedItem = closest > 0 ? closest : null;
            }
        }
        finally
        {
            _isUpdatingFontControls = false;
        }
    }

    private void FontSizeComboBox_Loaded(object sender, RoutedEventArgs e)
    {
        AttachFontSizeTextBoxHandlers();
    }

    private void AttachFontSizeTextBoxHandlers()
    {
        if (FontSizeComboBox == null || _fontSizeEditableTextBox != null)
            return;

        FontSizeComboBox.ApplyTemplate();
        _fontSizeEditableTextBox = FontSizeComboBox.Template?
            .FindName("PART_EditableTextBox", FontSizeComboBox) as WpfTextBox;

        if (_fontSizeEditableTextBox != null)
        {
            _fontSizeEditableTextBox.PreviewTextInput += FontSizeTextBox_PreviewTextInput;
            System.Windows.DataObject.AddPastingHandler(_fontSizeEditableTextBox, FontSizeTextBox_Pasting);
        }
    }

    private void FontSizeTextBox_PreviewTextInput(object sender, TextCompositionEventArgs e)
    {
        if (_fontSizeEditableTextBox == null)
            return;

        string prospectiveText = GetProspectiveFontSizeText(e.Text);
        e.Handled = !IsPartialFontSizeTextValid(prospectiveText);
    }

    private void FontSizeTextBox_Pasting(object sender, DataObjectPastingEventArgs e)
    {
        if (_fontSizeEditableTextBox == null)
            return;

        if (e.DataObject.GetData(System.Windows.DataFormats.Text) is string pastedText)
        {
            string prospectiveText = GetProspectiveFontSizeText(pastedText);
            if (!IsPartialFontSizeTextValid(prospectiveText))
            {
                e.CancelCommand();
            }
        }
        else
        {
            e.CancelCommand();
        }
    }

    private string GetProspectiveFontSizeText(string newText)
    {
        if (_fontSizeEditableTextBox == null)
            return FontSizeComboBox?.Text ?? string.Empty;

        int selectionStart = _fontSizeEditableTextBox.SelectionStart;
        int selectionLength = _fontSizeEditableTextBox.SelectionLength;
        string currentText = _fontSizeEditableTextBox.Text ?? string.Empty;

        return currentText.Remove(selectionStart, selectionLength).Insert(selectionStart, newText);
    }

    private bool IsPartialFontSizeTextValid(string text)
    {
        if (string.IsNullOrEmpty(text))
            return true;

        if (string.IsNullOrWhiteSpace(text))
            return false;

        string decimalSeparator = CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator;

        if (text.Equals(decimalSeparator, StringComparison.Ordinal))
            return true;

        if (text.EndsWith(decimalSeparator, StringComparison.Ordinal))
        {
            string withoutSeparator = text[..^decimalSeparator.Length];
            if (string.IsNullOrEmpty(withoutSeparator))
                return true;

            return double.TryParse(withoutSeparator, NumberStyles.Number, CultureInfo.CurrentCulture, out _);
        }

        return double.TryParse(text, NumberStyles.Number, CultureInfo.CurrentCulture, out _);
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
        AttachFontSizeTextBoxHandlers();

        if (e.Key == Key.Enter)
        {
            ApplyFontSizeFromTextInput();
            e.Handled = true;
            return;
        }

        if (!IsFontSizeKeyAllowed(e))
        {
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
            FontSizeComboBox.Text = ClampFontSizeInPoints(size).ToString("0.#");
            MarkDirty();
        }
        else
        {
            UpdateFontControlsFromSelection();
        }
    }

    private bool IsFontSizeKeyAllowed(KeyEventArgs e)
    {
        if ((Keyboard.Modifiers & (ModifierKeys.Control | ModifierKeys.Alt | ModifierKeys.Windows)) != 0)
            return true;

        switch (e.Key)
        {
            case Key.Back:
            case Key.Delete:
            case Key.Tab:
            case Key.Left:
            case Key.Right:
            case Key.Up:
            case Key.Down:
            case Key.Home:
            case Key.End:
            case Key.PageUp:
            case Key.PageDown:
            case Key.Escape:
                return true;
        }

        if (e.Key >= Key.F1 && e.Key <= Key.F24)
            return true;

        if (e.Key >= Key.D0 && e.Key <= Key.D9 && (Keyboard.Modifiers & ModifierKeys.Shift) == 0)
            return true;

        if (e.Key >= Key.NumPad0 && e.Key <= Key.NumPad9)
            return true;

        string decimalSeparator = CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator;

        if (IsDecimalKey(e.Key, decimalSeparator))
            return IsPartialFontSizeTextValid(GetProspectiveFontSizeText(decimalSeparator));

        return false;
    }

    private static bool IsDecimalKey(Key key, string decimalSeparator)
    {
        if (key == Key.Decimal)
            return true;

        if (decimalSeparator == "." && key == Key.OemPeriod)
            return true;

        if (decimalSeparator == "," && key == Key.OemComma)
            return true;

        return false;
    }

    private FontFamily? GetSelectionOrCaretFontFamily(TextSelection selection)
    {
        if (EditorBox == null)
            return null;

        object selectionFamily = selection.GetPropertyValue(Inline.FontFamilyProperty);
        if (selectionFamily is FontFamily fontFamily)
            return fontFamily;

        if (selection.IsEmpty)
        {
            TextPointer? caretPosition = EditorBox.CaretPosition;
            if (caretPosition != null)
            {
                object caretFamily = new TextRange(caretPosition, caretPosition)
                    .GetPropertyValue(Inline.FontFamilyProperty);
                if (caretFamily is FontFamily caretFontFamily)
                    return caretFontFamily;
            }
        }

        return EditorBox.Document?.FontFamily;
    }

    private FontFamily? ResolveComboFontFamily(FontFamily fontFamily)
    {
        if (FontFamilyComboBox == null)
            return null;

        var comboFamilies = FontFamilyComboBox.Items.Cast<FontFamily>();
        FontFamily? directMatch = comboFamilies
            .FirstOrDefault(f => string.Equals(f.Source, fontFamily.Source, StringComparison.OrdinalIgnoreCase));
        if (directMatch != null)
            return directMatch;

        foreach (string candidate in EnumerateFontFamilyCandidates(fontFamily.Source))
        {
            FontFamily? candidateMatch = comboFamilies
                .FirstOrDefault(f => string.Equals(f.Source, candidate, StringComparison.OrdinalIgnoreCase));
            if (candidateMatch != null)
                return candidateMatch;
        }

        foreach (string candidate in EnumerateFontFamilyCandidates(_styleConfiguration.BodyFontFamily))
        {
            FontFamily? candidateMatch = comboFamilies
                .FirstOrDefault(f => string.Equals(f.Source, candidate, StringComparison.OrdinalIgnoreCase));
            if (candidateMatch != null)
                return candidateMatch;
        }

        return null;
    }

    private static IEnumerable<string> EnumerateFontFamilyCandidates(string fontFamilySource)
    {
        return fontFamilySource
            .Split(',', StringSplitOptions.RemoveEmptyEntries)
            .Select(f => f.Trim().Trim('\'', '\"'))
            .Where(f => !string.IsNullOrWhiteSpace(f));
    }

    private Color? ShowColorPicker(Color? initialColor)
    {
        using var dialog = new System.Windows.Forms.ColorDialog
        {
            FullOpen = true
        };

        if (initialColor.HasValue)
        {
            dialog.Color = DrawingColor.FromArgb(
                initialColor.Value.A,
                initialColor.Value.R,
                initialColor.Value.G,
                initialColor.Value.B);
        }

        System.Windows.Forms.DialogResult result = dialog.ShowDialog();
        if (result != System.Windows.Forms.DialogResult.OK)
            return null;

        return Color.FromArgb(dialog.Color.A, dialog.Color.R, dialog.Color.G, dialog.Color.B);
    }

    private Brush CloneAndFreezeBrush(Brush brush)
    {
        if (brush.IsFrozen)
            return brush;

        Brush clone = brush.Clone();
        if (clone.CanFreeze)
            clone.Freeze();

        return clone;
    }

    private void ApplyBrushToSelection(DependencyProperty property, Brush? brush)
    {
        if (EditorBox == null || !CanFormat())
            return;

        object value = brush == null ? DependencyProperty.UnsetValue : CloneAndFreezeBrush(brush);
        EditorBox.Selection.ApplyPropertyValue(property, value);
    }

    private void ApplyHighlightBrush(Brush? brush)
    {
        if (EditorBox == null || !CanFormat())
            return;

        ApplyBrushToSelection(TextElement.BackgroundProperty, brush);
        MarkDirty();
        UpdateFormattingControls();
    }

    private void ApplyTextColorBrush(Brush? brush)
    {
        if (EditorBox == null || !CanFormat())
            return;

        ApplyBrushToSelection(Inline.ForegroundProperty, brush);
        MarkDirty();
        UpdateFormattingControls();
    }

    private void ClearSelectionColors()
    {
        if (!CanFormat())
            return;

        if (EditorBox == null)
            return;

        EditorBox.BeginChange();
        try
        {
            ApplyBrushToSelection(TextElement.BackgroundProperty, null);
            ApplyBrushToSelection(Inline.ForegroundProperty, null);
        }
        finally
        {
            EditorBox.EndChange();
        }

        MarkDirty();
        UpdateFormattingControls();
    }

    private void ApplyForeground(Color color)
    {
        ApplyTextColorBrush(new SolidColorBrush(color));
    }

    private void ApplyHighlight(Color color)
    {
        ApplyHighlightBrush(new SolidColorBrush(color));
    }

    private void ShowHighlightColorDialog()
    {
        Color? current = null;
        if (EditorBox != null)
        {
            object property = EditorBox.Selection.GetPropertyValue(TextElement.BackgroundProperty);
            if (property is SolidColorBrush brush)
                current = brush.Color;
        }

        Color? picked = ShowColorPicker(current);
        if (picked.HasValue)
            ApplyHighlight(picked.Value);
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

        UpdateFormattingControls();
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

        UpdateFormattingControls();
    }

    private void ToggleStrikethrough()
    {
        if (EditorBox == null || !CanFormat())
            return;

        TextSelection selection = EditorBox.Selection;
        object currentDecorations = selection.GetPropertyValue(Inline.TextDecorationsProperty);

        var current = currentDecorations as TextDecorationCollection;
        bool hasStrikethrough = current != null && current.Any(d => d.Location == TextDecorationLocation.Strikethrough);

        if (hasStrikethrough)
        {
            var updated = new TextDecorationCollection(current ?? new TextDecorationCollection());
            foreach (var decoration in TextDecorations.Strikethrough)
            {
                updated.Remove(decoration);
            }

            selection.ApplyPropertyValue(Inline.TextDecorationsProperty, updated.Count > 0 ? updated : null);
        }
        else
        {
            var updated = new TextDecorationCollection(current ?? new TextDecorationCollection())
            {
                TextDecorations.Strikethrough[0]
            };
            selection.ApplyPropertyValue(Inline.TextDecorationsProperty, updated);
        }

        UpdateFormattingControls();
    }

    private void ToggleMonospaced()
    {
        if (EditorBox == null || !CanFormat())
            return;

        // The UI command flips the font family between the body stack and the mono stack.  This
        // is the single toggle used by the toolbar, keyboard shortcut (Ctrl+Shift+M), and context
        // menu; downstream exports and ODT serialization read the IsMonospaced flag set here.
        TextSelection selection = EditorBox.Selection;
        object currentFamily = selection.GetPropertyValue(Inline.FontFamilyProperty);
        bool isMonospaced = currentFamily != DependencyProperty.UnsetValue
            && currentFamily is FontFamily family
            && IsMonospacedFont(family);

        EditorBox.BeginChange();
        try
        {
            selection.ApplyPropertyValue(
                Inline.FontFamilyProperty,
                isMonospaced ? GetBodyFontFamily() : GetMonoFontFamily());
        }
        finally
        {
            EditorBox.EndChange();
        }

        UpdateFormattingControls();
    }

    private void ToggleMonospacedFromUi()
    {
        if (!CanFormat())
            return;

        ToggleMonospaced();
        MarkDirty();
    }

    private void ToggleItalic()
    {
        if (EditorBox == null || !CanFormat())
            return;

        TextSelection selection = EditorBox.Selection;
        object currentStyle = selection.GetPropertyValue(Inline.FontStyleProperty);

        if (currentStyle != DependencyProperty.UnsetValue && currentStyle.Equals(System.Windows.FontStyles.Italic))
        {
            selection.ApplyPropertyValue(Inline.FontStyleProperty, System.Windows.FontStyles.Normal);
        }
        else
        {
            selection.ApplyPropertyValue(Inline.FontStyleProperty, System.Windows.FontStyles.Italic);
        }

        UpdateFormattingControls();
    }

    private void ToggleBulletedList()
    {
        ApplyListStyle(TextMarkerStyle.Disc);
    }

    private void ToggleNumberedList()
    {
        ApplyListStyle(TextMarkerStyle.Decimal);
    }

    private void ToggleLetteredList()
    {
        ApplyListStyle(TextMarkerStyle.LowerLatin);
    }

    private void ClearListFormatting()
    {
        RemoveListFormattingFromSelection();
    }

    private void BulletedListButton_Click(object sender, RoutedEventArgs e)
    {
        if (!CanFormat())
            return;

        ToggleBulletedList();
    }

    private void BulletedListMenuItem_Click(object sender, RoutedEventArgs e)
    {
        BulletedListButton_Click(sender, e);
    }

    private void NumberedListButton_Click(object sender, RoutedEventArgs e)
    {
        if (!CanFormat())
            return;

        ToggleNumberedList();
    }

    private void NumberedListMenuItem_Click(object sender, RoutedEventArgs e)
    {
        NumberedListButton_Click(sender, e);
    }

    private void LetteredListButton_Click(object sender, RoutedEventArgs e)
    {
        if (!CanFormat())
            return;

        ToggleLetteredList();
    }

    private void LetteredListMenuItem_Click(object sender, RoutedEventArgs e)
    {
        LetteredListButton_Click(sender, e);
    }

    private void ClearListFormattingButton_Click(object sender, RoutedEventArgs e)
    {
        if (!CanFormat())
            return;

        ClearListFormatting();
    }

    private void ClearListFormattingMenuItem_Click(object sender, RoutedEventArgs e)
    {
        ClearListFormattingButton_Click(sender, e);
    }

    private void HighlightButton_Click(object sender, RoutedEventArgs e)
    {
        if (!CanFormat())
            return;

        if (HighlightButton?.ContextMenu != null)
        {
            HighlightButton.ContextMenu.PlacementTarget = HighlightButton;
            HighlightButton.ContextMenu.Placement = PlacementMode.Bottom;
            HighlightButton.ContextMenu.IsOpen = true;
            return;
        }

        ShowHighlightColorDialog();
    }

    private void HighlightToolbarColor_Click(object sender, RoutedEventArgs e)
    {
        if (!CanFormat())
            return;

        if (sender is System.Windows.Controls.Control control)
        {
            if (control.Tag is string tag && string.Equals(tag, "None", StringComparison.OrdinalIgnoreCase))
            {
                ApplyHighlightBrush(null);
                return;
            }

            Brush? brush = control.Background;
            if (brush == null && control.Tag is string colorText)
            {
                object? converted = System.Windows.Media.ColorConverter.ConvertFromString(colorText);
                if (converted is Color color)
                    brush = new SolidColorBrush(color);
            }

            if (brush != null)
                ApplyHighlightBrush(brush);
        }
    }

    private void HighlightMoreColorsMenuItem_Click(object sender, RoutedEventArgs e)
    {
        if (!CanFormat())
            return;

        ShowHighlightColorDialog();
    }

    private void ClearHighlightMenuItem_Click(object sender, RoutedEventArgs e)
    {
        if (!CanFormat())
            return;

        ApplyHighlightBrush(null);
    }

    private void TextColorButton_Click(object sender, RoutedEventArgs e)
    {
        if (!CanFormat())
            return;

        Color? current = null;
        if (EditorBox != null)
        {
            object property = EditorBox.Selection.GetPropertyValue(TextElement.ForegroundProperty);
            if (property is SolidColorBrush brush)
                current = brush.Color;
        }

        Color? picked = ShowColorPicker(current);
        if (picked.HasValue)
            ApplyForeground(picked.Value);
    }

    private void TextColorToolbarColor_Click(object sender, RoutedEventArgs e)
    {
        if (!CanFormat())
            return;

        if (sender is System.Windows.Controls.Control control && control.Background is Brush brush)
            ApplyTextColorBrush(brush);
    }

    private void ClearToolbarColors_Click(object sender, RoutedEventArgs e)
    {
        ClearSelectionColors();
    }

    private bool CanFormat()
    {
        return _currentDocumentType == DocumentType.RichText
            || _currentDocumentType == DocumentType.OpenDocument;
    }

    private bool IsSelectionInList()
    {
        if (EditorBox == null)
            return false;

        return GetListItemsFromSelection().Count > 0;
    }

    private void UpdateFormattingControls()
    {
        bool canFormat = CanFormat();

        if (BoldButton != null)
        {
            BoldButton.IsEnabled = canFormat;
            BoldButton.IsChecked = canFormat && IsSelectionBold();
        }
        if (ItalicButton != null)
        {
            ItalicButton.IsEnabled = canFormat;
            ItalicButton.IsChecked = canFormat && IsSelectionItalic();
        }
        if (UnderlineButton != null)
        {
            UnderlineButton.IsEnabled = canFormat;
            UnderlineButton.IsChecked = canFormat && IsSelectionUnderlined();
        }
        if (StrikethroughButton != null)
        {
            StrikethroughButton.IsEnabled = canFormat;
            StrikethroughButton.IsChecked = canFormat && IsSelectionStruckThrough();
        }
        if (MonospacedButton != null)
        {
            MonospacedButton.IsEnabled = canFormat;
            MonospacedButton.IsChecked = canFormat && IsSelectionMonospaced();
        }
        if (HighlightButton != null)
            HighlightButton.IsEnabled = canFormat;
        if (TextColorButton != null)
            TextColorButton.IsEnabled = canFormat;
        if (NumberedListMenuItem != null)
            NumberedListMenuItem.IsEnabled = canFormat;
        if (LetteredListMenuItem != null)
            LetteredListMenuItem.IsEnabled = canFormat;
        if (BulletedListMenuItem != null)
            BulletedListMenuItem.IsEnabled = canFormat;
        if (ClearListFormattingMenuItem != null)
            ClearListFormattingMenuItem.IsEnabled = canFormat && IsSelectionInList();
        if (AlignLeftButton != null)
            AlignLeftButton.IsEnabled = canFormat;
        if (AlignCenterButton != null)
            AlignCenterButton.IsEnabled = canFormat;
        if (AlignRightButton != null)
            AlignRightButton.IsEnabled = canFormat;

        UpdateFontControlsFromSelection();
    }

    private bool IsSelectionBold()
    {
        if (EditorBox == null)
            return false;

        object weight = EditorBox.Selection.GetPropertyValue(Inline.FontWeightProperty);
        return weight is FontWeight fontWeight && fontWeight == FontWeights.Bold;
    }

    private bool IsSelectionItalic()
    {
        if (EditorBox == null)
            return false;

        object style = EditorBox.Selection.GetPropertyValue(Inline.FontStyleProperty);
        return style is System.Windows.FontStyle fontStyle && fontStyle == System.Windows.FontStyles.Italic;
    }

    private bool IsSelectionUnderlined()
    {
        if (EditorBox == null)
            return false;

        var decorations = EditorBox.Selection.GetPropertyValue(Inline.TextDecorationsProperty)
            as TextDecorationCollection;
        return decorations != null && decorations.Contains(TextDecorations.Underline[0]);
    }

    private bool IsSelectionStruckThrough()
    {
        if (EditorBox == null)
            return false;

        var decorations = EditorBox.Selection.GetPropertyValue(Inline.TextDecorationsProperty)
            as TextDecorationCollection;
        return decorations != null && decorations.Any(d => d.Location == TextDecorationLocation.Strikethrough);
    }

    private bool IsSelectionMonospaced()
    {
        if (EditorBox == null)
            return false;

        object family = EditorBox.Selection.GetPropertyValue(Inline.FontFamilyProperty);
        return family is FontFamily fontFamily && IsMonospacedFont(fontFamily);
    }

    private void AlignLeftButton_Click(object sender, RoutedEventArgs e)
    {
        SetTextAlignment(TextAlignment.Left);
    }

    private void AlignCenterButton_Click(object sender, RoutedEventArgs e)
    {
        SetTextAlignment(TextAlignment.Center);
    }

    private void AlignRightButton_Click(object sender, RoutedEventArgs e)
    {
        SetTextAlignment(TextAlignment.Right);
    }

    private void SetTextAlignment(TextAlignment alignment)
    {
        if (EditorBox == null || !CanFormat())
            return;

        EditorBox.Selection.ApplyPropertyValue(Paragraph.TextAlignmentProperty, alignment);
        MarkDirty();
    }
}

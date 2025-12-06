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
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Threading;
using System.Xml.Linq;
using Microsoft.Win32;
using Markdig;
using PdfSharp.Drawing;
using PdfSharp.Pdf;

namespace MGPad;

public enum DocumentType
{
    RichText,
    PlainText,
    Markdown,
    OpenDocument
}

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
    private static readonly Regex NumberRegex = new(
        "[-+]?(?:\\d+\\.?\\d*|\\.\\d+)(?:[eE][-+]?\\d+)?",
        RegexOptions.Compiled);
    private readonly MarkdownPipeline _markdownPipeline;
    private readonly AutosaveService _autosaveService;
    private Guid _untitledDocumentId = Guid.NewGuid();

    private static readonly TimeSpan AutosaveInterval = TimeSpan.FromSeconds(60);

    private sealed class StyleConfiguration
    {
        // Body and mono font stacks prefer Windows defaults while including explicit Japanese
        // fallbacks.  The comma-separated list follows the same syntax as CSS font-family:
        // choose the first installed face and fall back to later entries when characters (e.g.,
        // CJK glyphs) are missing.  This keeps Unicode-heavy notes readable without requiring
        // the exact font set from the authoring machine.
        public string BodyFontFamily { get; init; } = "Segoe UI, 'Yu Gothic UI'";

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

    private const string HelpText =
        "MGPad quick help:\n" +
        "\n" +
        "• Ctrl+N / Ctrl+O / Ctrl+S to create, open, or save documents.\n" +
        "• Ctrl+B / Ctrl+I / Ctrl+U / Ctrl+Shift+X to toggle bold, italic, underline, or strikethrough.\n" +
        "• Alt+Shift+T inserts a timestamp.\n" +
        "• Use View → Markdown Mode to preview Markdown on the right.\n" +
        "• Toggle JP/EN switches input language when both are available.\n" +
        "• Use the zoom controls in the status bar to adjust text size.";

    public MainWindow()
    {
        InitializeComponent();
        _recentDocumentsFilePath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "MGPad",
            "recent-documents.txt");
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
            EditorBox.CommandBindings.Add(new CommandBinding(
                ToggleMonospacedCommand,
                (s, e) =>
                {
                    ToggleMonospacedFromUi();
                    e.Handled = true;
                },
                (s, e) => e.CanExecute = CanFormat()));
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
    }

    private void HelpMenuItem_Click(object sender, RoutedEventArgs e)
    {
        MessageBox.Show(
            HelpText,
            "MGPad Help",
            MessageBoxButton.OK,
            MessageBoxImage.Information);
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
        if (!_isMarkdownMode || MarkdownPreviewBrowser == null)
            return;

        string markdown = GetEditorPlainText();
        string normalized = markdown.Replace("\r\n", "\n").Replace("\r", "\n");

        string html = string.IsNullOrWhiteSpace(normalized)
            ? string.Empty
            : Markdown.ToHtml(normalized, _markdownPipeline);

        string background = _isNightMode ? "#2d2d30" : "#ffffff";
        string foreground = _isNightMode ? "#f5f5f5" : "#000000";

        string document = $"<!DOCTYPE html><html><head><meta charset=\"utf-8\" /><style>body {{ font-family: 'Segoe UI', 'Yu Gothic UI', sans-serif; padding: 12px; color: {foreground}; background: {background}; }} code, pre {{ font-family: 'Cascadia Code', Consolas, 'Courier New', monospace; }} pre {{ background: rgba(0,0,0,0.04); padding: 8px; overflow-x: auto; }} a {{ color: #0066cc; }}</style></head><body>{html}</body></html>";

        MarkdownPreviewBrowser.NavigateToString(document);
    }

    private FlowDocument CreateStyledDocument()
    {
        return new FlowDocument
        {
            FontFamily = GetBodyFontFamily()
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
    }

    private void LoadRtfIntoEditor(string path)
    {
        if (EditorBox == null)
            return;

        EditorBox.Document = CreateStyledDocument();
        var range = new TextRange(EditorBox.Document.ContentStart, EditorBox.Document.ContentEnd);

        using var stream = new FileStream(path, FileMode.Open, FileAccess.Read);
        range.Load(stream, DataFormats.Rtf);
        ApplyTheme();
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
                    wpfRun.FontStyle = FontStyles.Italic;
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

        string timestamp = DateTime.Now.ToString("dddd, MMMM d, yyyy h:mm tt");

        TextRange selection = EditorBox.Selection;

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

    private void SumSelectionMenuItem_Click(object sender, RoutedEventArgs e)
    {
        if (EditorBox == null)
            return;

        string selectedText = EditorBox.Selection.Text;
        if (string.IsNullOrWhiteSpace(selectedText))
        {
            MessageBox.Show(
                "Select one or more numbers to sum.",
                "Sum selection",
                MessageBoxButton.OK,
                MessageBoxImage.Information);
            return;
        }

        var numbers = ExtractNumbersFromSelection(selectedText);
        if (numbers.Count == 0)
        {
            MessageBox.Show(
                "No numbers were found in the selection.",
                "Sum selection",
                MessageBoxButton.OK,
                MessageBoxImage.Information);
            return;
        }

        double sum = numbers.Sum();
        string sumText = sum.ToString(CultureInfo.InvariantCulture);
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

    private static List<double> ExtractNumbersFromSelection(string selectedText)
    {
        var numbers = new List<double>();

        foreach (Match match in NumberRegex.Matches(selectedText))
        {
            if (double.TryParse(match.Value, NumberStyles.Float, CultureInfo.InvariantCulture, out double value))
                numbers.Add(value);
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

    private void SetCurrentFile(string? path, DocumentType type)
    {
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

            MessageBox.Show(this,
                "ODT export completed:\n" + odtPath,
                "Export as ODT",
                MessageBoxButton.OK,
                MessageBoxImage.Information);
        }
        catch (Exception ex)
        {
            MessageBox.Show(this,
                "Failed to export ODT.\n\n" + ex.Message,
                "Export as ODT",
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
                bool isItalic = run.FontStyle == FontStyles.Italic;
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
        var fontCache = new Dictionary<(string Family, XFontStyle Style), XFont>();
        XFont regularFont = GetOrCreateFont(
            fontCache,
            _styleConfiguration.BodyFontFamily,
            fontSize,
            XFontStyle.Regular);

        // Layout: 1-inch margins (72 points per inch)
        double marginLeft = 72;
        double marginTop = 72;
        double marginRight = 72;
        double marginBottom = 72;

        double y = marginTop;
        foreach (var paragraph in paragraphs)
        {
            double usableWidth = page.Width - marginLeft - marginRight;

            var wrappedLines = WrapParagraphRuns(
                paragraph.Runs,
                gfx,
                usableWidth,
                run => GetFontForRun(run, fontCache, _styleConfiguration, fontSize));

            double lastLineHeight = 0;

            foreach (var line in wrappedLines)
            {
                double lineHeight = GetLineHeight(line, gfx, lineSpacingFactor, regularFont);

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
            double spanHeight = span.Font.GetHeight(gfx);
            maxHeight = Math.Max(maxHeight, spanHeight);
        }

        double baseHeight = maxHeight > 0
            ? maxHeight
            : fallbackFont.GetHeight(gfx);

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
        Dictionary<(string Family, XFontStyle Style), XFont> fontCache,
        StyleConfiguration styleConfiguration,
        double fontSize)
    {
        XFontStyle style = XFontStyle.Regular;

        if (run.IsBold)
            style |= XFontStyle.Bold;
        if (run.IsItalic)
            style |= XFontStyle.Italic;
        if (run.IsUnderline)
            style |= XFontStyle.Underline;
        if (run.IsStrikethrough)
            style |= XFontStyle.Strikeout;

        string selectedFamily = GetFontFamilyForRun(run, styleConfiguration);

        return GetOrCreateFont(fontCache, selectedFamily, fontSize, style);
    }

    private static XFont GetOrCreateFont(
        Dictionary<(string Family, XFontStyle Style), XFont> fontCache,
        string fontFamily,
        double fontSize,
        XFontStyle style)
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
        var contentDoc = new XDocument(
            new XDeclaration("1.0", "UTF-8", "yes"),
            new XElement(office + "document-content",
                new XAttribute(XNamespace.Xmlns + "office", office),
                new XAttribute(XNamespace.Xmlns + "style", style),
                new XAttribute(XNamespace.Xmlns + "text", text),
                new XAttribute(XNamespace.Xmlns + "fo", fo),
                new XElement(office + "automatic-styles", automaticStyles),
                new XElement(office + "body",
                    new XElement(office + "text",
                        paragraphs.Select(p => CreateOdtParagraph(p, automaticStyles, text, style, fo))))));

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

    private XElement CreateOdtParagraph(
        PdfParagraph paragraph,
        List<XElement> automaticStyles,
        XNamespace text,
        XNamespace style,
        XNamespace fo)
    {
        var paragraphElement = new XElement(text + "p");

        foreach (var run in paragraph.Runs)
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
            MessageBox.Show($"Failed to load document: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
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
                    range.Save(stream, DataFormats.Rtf);
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
                MessageBox.Show($"Failed to save document: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
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

        var dialog = new SaveFileDialog
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

    private void EditorBox_PreviewKeyDown(object sender, KeyEventArgs e)
    {
        if (EditorBox == null || !CanFormat())
            return;

        if (e.Key == Key.Tab && IsSelectionInList())
        {
            if ((Keyboard.Modifiers & ModifierKeys.Shift) == ModifierKeys.Shift)
            {
                EditingCommands.DecreaseIndentation.Execute(null, EditorBox);
            }
            else
            {
                EditingCommands.IncreaseIndentation.Execute(null, EditorBox);
            }

            e.Handled = true;
            MarkDirty();
        }
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

        if (currentStyle != DependencyProperty.UnsetValue && currentStyle.Equals(FontStyles.Italic))
        {
            selection.ApplyPropertyValue(Inline.FontStyleProperty, FontStyles.Normal);
        }
        else
        {
            selection.ApplyPropertyValue(Inline.FontStyleProperty, FontStyles.Italic);
        }

        UpdateFormattingControls();
    }

    private void ToggleBulletedList()
    {
        if (EditorBox == null || !CanFormat())
            return;

        EditingCommands.ToggleBullets.Execute(null, EditorBox);
    }

    private void ToggleNumberedList()
    {
        if (EditorBox == null || !CanFormat())
            return;

        EditingCommands.ToggleNumbering.Execute(null, EditorBox);
    }

    private void BoldButton_Click(object sender, RoutedEventArgs e)
    {
        if (!CanFormat())
            return;

        ToggleBold();
        MarkDirty();
    }

    private void ItalicButton_Click(object sender, RoutedEventArgs e)
    {
        if (!CanFormat())
            return;

        ToggleItalic();
        MarkDirty();
    }

    private void UnderlineButton_Click(object sender, RoutedEventArgs e)
    {
        if (!CanFormat())
            return;

        ToggleUnderline();
        MarkDirty();
    }

    private void StrikethroughButton_Click(object sender, RoutedEventArgs e)
    {
        if (!CanFormat())
            return;

        ToggleStrikethrough();
        MarkDirty();
    }

    private void MonospacedButton_Click(object sender, RoutedEventArgs e)
    {
        ToggleMonospacedFromUi();
    }

    private void BulletedListButton_Click(object sender, RoutedEventArgs e)
    {
        if (!CanFormat())
            return;

        ToggleBulletedList();
        MarkDirty();
    }

    private void NumberedListButton_Click(object sender, RoutedEventArgs e)
    {
        if (!CanFormat())
            return;

        ToggleNumberedList();
        MarkDirty();
    }

    private bool CanFormat()
    {
        return _currentDocumentType == DocumentType.RichText
            || _currentDocumentType == DocumentType.OpenDocument;
    }

    private System.Windows.Documents.List? GetAncestorList(TextPointer position)
    {
        DependencyObject? current = position.Parent;

        while (current != null && current is not FlowDocument)
        {
            if (current is System.Windows.Documents.List list)
                return list;

            current = LogicalTreeHelper.GetParent(current);
        }

        return null;
    }

    private bool IsSelectionInList()
    {
        if (EditorBox == null)
            return false;

        return GetAncestorList(EditorBox.Selection.Start) != null
            || GetAncestorList(EditorBox.Selection.End) != null;
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
        if (BulletedListButton != null)
            BulletedListButton.IsEnabled = canFormat;
        if (NumberedListButton != null)
            NumberedListButton.IsEnabled = canFormat;
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
        return style is FontStyle fontStyle && fontStyle == FontStyles.Italic;
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

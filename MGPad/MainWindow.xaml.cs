using System;
using System.ComponentModel;
using System.IO;
using System.Globalization;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Input;
using System.Text;
using Microsoft.Win32;
using System.Windows.Threading;

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

    private string? _currentFilePath;
    private DocumentType _currentDocumentType;
    private bool _isDirty;
    private bool _isLoadingDocument;
    private bool _allowCloseWithoutPrompt;
    private bool _isMarkdownMode = false;
    private readonly DispatcherTimer _markdownPreviewTimer;
    private CultureInfo? _englishInputLanguage;
    private CultureInfo? _japaneseInputLanguage;

    public MainWindow()
    {
        InitializeComponent();
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
        InitializePreferredInputLanguages();

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
    }

    private void LoadRtfIntoEditor(string path)
    {
        if (EditorBox == null)
            return;

        EditorBox.Document = new FlowDocument();
        var range = new TextRange(EditorBox.Document.ContentStart, EditorBox.Document.ContentEnd);

        using var stream = new FileStream(path, FileMode.Open, FileAccess.Read);
        range.Load(stream, DataFormats.Rtf);
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
            LoadDocumentFromFile(dialog.FileName);
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
    }
}

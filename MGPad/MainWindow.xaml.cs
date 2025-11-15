using System;
using System.IO;
using System.Windows;
using System.Windows.Documents;

namespace MGPad;

public enum DocumentType
{
    RichText,
    PlainText,
    Markdown
}

public partial class MainWindow : Window
{
    private string? _currentFilePath;
    private DocumentType _currentDocumentType;
    private bool _isDirty;

    public MainWindow()
    {
        InitializeComponent();
        InitializeDocumentState();
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
        var fileName = string.IsNullOrEmpty(_currentFilePath)
            ? "Untitled"
            : Path.GetFileName(_currentFilePath);

        var dirtyMarker = _isDirty ? "* " : string.Empty;
        Title = $"MGPad - {dirtyMarker}{fileName}";
    }

    private void FileNew_Click(object sender, RoutedEventArgs e)
    {
        EditorBox.Document = new FlowDocument();
        SetCurrentFile(null, DocumentType.RichText);
        MarkClean();
    }

    private void FileOpen_Click(object sender, RoutedEventArgs e)
    {
    }

    private void FileSave_Click(object sender, RoutedEventArgs e)
    {
    }

    private void FileSaveAs_Click(object sender, RoutedEventArgs e)
    {
    }

    private void FileExit_Click(object sender, RoutedEventArgs e)
    {
    }

    private void LoadDocumentFromFile(string path)
    {
        try
        {
            var documentType = DetermineDocumentType(path);
            EditorBox.Document = new FlowDocument();
            var range = new TextRange(EditorBox.Document.ContentStart, EditorBox.Document.ContentEnd);

            if (documentType == DocumentType.RichText)
            {
                using var stream = new FileStream(path, FileMode.Open, FileAccess.Read);
                range.Load(stream, DataFormats.Rtf);
            }
            else
            {
                var text = File.ReadAllText(path);
                range.Text = text;
            }

            SetCurrentFile(path, documentType);
            MarkClean();
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Failed to load document: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void SaveDocumentToFile(string path)
    {
        try
        {
            var documentType = DetermineDocumentType(path);
            var range = new TextRange(EditorBox.Document.ContentStart, EditorBox.Document.ContentEnd);

            if (documentType == DocumentType.RichText)
            {
                using var stream = new FileStream(path, FileMode.Create, FileAccess.Write);
                range.Save(stream, DataFormats.Rtf);
            }
            else
            {
                File.WriteAllText(path, range.Text);
            }

            SetCurrentFile(path, documentType);
            MarkClean();
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Failed to save document: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
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
}

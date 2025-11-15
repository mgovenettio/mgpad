using System.IO;
using System.Windows;

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
}

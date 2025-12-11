using System;
using System.Diagnostics;
using System.IO;
using System.Text.Json;
using System.Threading.Tasks;
using System.Windows.Threading;

namespace MGPad;

internal sealed record AutosaveContext(string BaseName, DocumentType DocumentType, string Extension, string? OriginalPath, bool IsUntitled);

internal sealed class AutosaveService
{
    private readonly DispatcherTimer _timer;
    private readonly Func<bool> _isDirty;
    private readonly Func<AutosaveContext> _contextProvider;
    private readonly Func<string, DocumentType, bool> _writeCallback;
    private readonly string _autosaveDirectory;
    private string? _lastAutosavePath;
    private string? _lastMetadataPath;
    private bool _isAutosaving;

    public AutosaveService(
        TimeSpan interval,
        Func<bool> isDirty,
        Func<AutosaveContext> contextProvider,
        Func<string, DocumentType, bool> writeCallback)
    {
        _isDirty = isDirty;
        _contextProvider = contextProvider;
        _writeCallback = writeCallback;
        _autosaveDirectory = GetAutosaveDirectory();

        _timer = new DispatcherTimer { Interval = interval };
        _timer.Tick += OnTimerTick;
    }

    public void Start() => _timer.Start();

    public void Stop() => _timer.Stop();

    public void OnManualSave()
    {
        DeleteCurrentAutosaveFiles();
    }

    public void DeleteCurrentAutosaveFiles()
    {
        TryDeleteFile(_lastAutosavePath);
        TryDeleteFile(_lastMetadataPath);
        _lastAutosavePath = null;
        _lastMetadataPath = null;
    }

    private async void OnTimerTick(object? sender, EventArgs e)
    {
        if (_isAutosaving)
        {
            return;
        }

        if (!_isDirty())
        {
            return;
        }

        var context = _contextProvider();
        if (string.IsNullOrWhiteSpace(context.BaseName))
        {
            return;
        }

        _isAutosaving = true;
        try
        {
            Directory.CreateDirectory(_autosaveDirectory);
            var autosavePath = Path.Combine(_autosaveDirectory, context.BaseName + context.Extension);
            var metadataPath = Path.Combine(_autosaveDirectory, context.BaseName + ".json");

            await Task.Run(() =>
            {
                try
                {
                    bool saved = false;
                    System.Windows.Application.Current?.Dispatcher.Invoke(() =>
                    {
                        saved = _writeCallback(autosavePath, context.DocumentType);
                    });

                    if (!saved)
                    {
                        return;
                    }

                    var metadata = new
                    {
                        OriginalPath = context.OriginalPath,
                        IsUntitled = context.IsUntitled,
                        TimestampUtc = DateTimeOffset.UtcNow
                    };

                    File.WriteAllText(metadataPath, JsonSerializer.Serialize(metadata, new JsonSerializerOptions
                    {
                        WriteIndented = true
                    }));

                    _lastAutosavePath = autosavePath;
                    _lastMetadataPath = metadataPath;
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"Autosave failed: {ex}");
                }
            });
        }
        finally
        {
            _isAutosaving = false;
        }
    }

    internal static string GetAutosaveDirectory()
    {
        return Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "MGPad",
            "Autosave");
    }

    internal static void TryDeleteFile(string? path)
    {
        if (string.IsNullOrWhiteSpace(path))
        {
            return;
        }

        try
        {
            if (File.Exists(path))
            {
                File.Delete(path);
            }
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"Failed to delete autosave file '{path}': {ex}");
        }
    }
}

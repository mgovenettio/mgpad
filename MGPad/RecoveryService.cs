using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text.Json;

namespace MGPad;

public sealed record RecoveryItem(
    string DisplayName,
    string? OriginalPath,
    DateTimeOffset TimestampUtc,
    string AutosavePath,
    string MetadataPath,
    DocumentType DocumentType,
    bool IsUntitled)
{
    public string TimestampLocal => TimestampUtc.ToLocalTime().ToString("g");
}

public sealed class RecoveryService
{
    private readonly string _autosaveDirectory;

    public RecoveryService()
    {
        _autosaveDirectory = AutosaveService.GetAutosaveDirectory();
    }

    public List<RecoveryItem> GetRecoverableItems()
    {
        var recoverable = new List<RecoveryItem>();

        if (!Directory.Exists(_autosaveDirectory))
        {
            return recoverable;
        }

        foreach (var metadataPath in Directory.GetFiles(_autosaveDirectory, "*.json"))
        {
            try
            {
                var baseName = Path.GetFileNameWithoutExtension(metadataPath);
                var autosavePath = Directory.GetFiles(_autosaveDirectory, baseName + ".*")
                    .FirstOrDefault(f => !string.Equals(Path.GetExtension(f), ".json", StringComparison.OrdinalIgnoreCase));

                if (string.IsNullOrWhiteSpace(autosavePath) || !File.Exists(autosavePath))
                {
                    continue;
                }

                var metadata = JsonSerializer.Deserialize<AutosaveMetadata>(File.ReadAllText(metadataPath));
                if (metadata is null)
                {
                    continue;
                }

                var displayName = string.IsNullOrWhiteSpace(metadata.OriginalPath)
                    ? "Untitled"
                    : Path.GetFileName(metadata.OriginalPath);

                recoverable.Add(new RecoveryItem(
                    displayName,
                    metadata.OriginalPath,
                    metadata.TimestampUtc,
                    autosavePath,
                    metadataPath,
                    DetermineDocumentType(autosavePath),
                    metadata.IsUntitled));
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Failed to parse autosave metadata '{metadataPath}': {ex}");
            }
        }

        return recoverable
            .OrderByDescending(item => item.TimestampUtc)
            .ToList();
    }

    public void Discard(RecoveryItem item)
    {
        AutosaveService.TryDeleteFile(item.AutosavePath);
        AutosaveService.TryDeleteFile(item.MetadataPath);
    }

    public void DiscardAll(IEnumerable<RecoveryItem> items)
    {
        foreach (var item in items)
        {
            Discard(item);
        }
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

    private sealed class AutosaveMetadata
    {
        public string? OriginalPath { get; set; }

        public bool IsUntitled { get; set; }

        public DateTimeOffset TimestampUtc { get; set; }
    }
}

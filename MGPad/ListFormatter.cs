using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Controls;
using System.Windows.Documents;

namespace MGPad;

internal static class ListFormatter
{
    private enum ListType
    {
        Numbered,
        Lettered
    }

    private sealed record LineInfo(TextPointer Start, TextPointer End, string Text)
    {
        public string NormalizedText => Text.TrimEnd('\r', '\n');
    }

    // Regexes capture indentation, the prefix marker (number or letter), punctuation, and spacing so we
    // can rebuild the marker without touching the rest of the line's content or formatting.
    private static readonly Regex NumberedListRegex = new(
        "^(?<prefix>(?<indent>\\s*)(?<marker>\\d+)(?<punct>[.)])(?<spacing>\\s+))",
        RegexOptions.Compiled);

    private static readonly Regex LetteredListRegex = new(
        "^(?<prefix>(?<indent>\\s*)(?<marker>[A-Za-z])(?<punct>[.)])(?<spacing>\\s+))",
        RegexOptions.Compiled);

    public static void RenumberLists(RichTextBox? editor)
    {
        if (editor?.Document is null)
            return;

        FlowDocument document = editor.Document;
        List<LineInfo> lines = GetLines(document).ToList();

        if (lines.Count == 0)
            return;

        SelectionSnapshot selectionSnapshot = CaptureSelection(editor);

        for (int i = 0; i < lines.Count; i++)
        {
            if (!TryGetListMatch(lines[i].NormalizedText, out ListType listType, out Match? firstMatch))
                continue;

            int blockStart = i;
            List<Match> blockMatches = new() { firstMatch };

            i++;
            while (i < lines.Count &&
                   TryGetListMatch(lines[i].NormalizedText, out ListType nextType, out Match? nextMatch) &&
                   nextType == listType)
            {
                blockMatches.Add(nextMatch);
                i++;
            }

            RenumberBlock(lines, blockStart, blockMatches, listType);
            i--;
        }

        RestoreSelection(editor, selectionSnapshot);
    }

    private static IEnumerable<LineInfo> GetLines(FlowDocument document)
    {
        TextPointer? lineStart = document.ContentStart.GetLineStartPosition(0);

        while (lineStart != null && lineStart.CompareTo(document.ContentEnd) < 0)
        {
            TextPointer? nextLineStart = lineStart.GetLineStartPosition(1);
            TextPointer lineEnd = nextLineStart ?? document.ContentEnd;
            string text = new TextRange(lineStart, lineEnd).Text;

            yield return new LineInfo(lineStart, lineEnd, text);

            if (nextLineStart == null)
                yield break;

            lineStart = nextLineStart;
        }
    }

    private static bool TryGetListMatch(string text, out ListType listType, out Match? match)
    {
        match = NumberedListRegex.Match(text);
        if (match.Success)
        {
            listType = ListType.Numbered;
            return true;
        }

        match = LetteredListRegex.Match(text);
        if (match.Success)
        {
            listType = ListType.Lettered;
            return true;
        }

        // Bullet markers intentionally fall through here so bullet lists remain untouched.
        listType = default;
        return false;
    }

    private static void RenumberBlock(IReadOnlyList<LineInfo> lines, int blockStart, IReadOnlyList<Match> matches, ListType listType)
    {
        bool useUppercaseLetters = listType == ListType.Lettered && char.IsUpper(matches[0].Groups["marker"].Value[0]);
        char letterBase = useUppercaseLetters ? 'A' : 'a';

        for (int i = 0; i < matches.Count; i++)
        {
            Match match = matches[i];
            LineInfo line = lines[blockStart + i];

            string indent = match.Groups["indent"].Value;
            string punctuation = match.Groups["punct"].Value;
            string spacing = match.Groups["spacing"].Value;

            string newMarker = listType switch
            {
                ListType.Numbered => (i + 1).ToString(),
                _ => BuildLetterMarker(letterBase, i)
            };

            string newPrefix = indent + newMarker + punctuation + spacing;
            string existingPrefix = match.Groups["prefix"].Value;

            if (newPrefix == existingPrefix)
                continue;

            TextPointer? prefixStart = line.Start;
            TextPointer? prefixEnd = prefixStart?.GetPositionAtOffset(existingPrefix.Length);

            if (prefixStart != null && prefixEnd != null)
                new TextRange(prefixStart, prefixEnd).Text = newPrefix;
        }
    }

    private static string BuildLetterMarker(char letterBase, int index)
    {
        int clamped = Math.Min(index, (letterBase == 'A' ? 'Z' : 'z') - letterBase);
        return ((char)(letterBase + clamped)).ToString();
    }

    private static (int lineIndex, int columnOffset) CapturePosition(RichTextBox editor, TextPointer position)
    {
        TextPointer lineStart = position.GetLineStartPosition(0) ?? position;

        int lineIndex = 0;
        TextPointer? walker = editor.Document.ContentStart.GetLineStartPosition(0);
        while (walker != null && walker.CompareTo(lineStart) < 0)
        {
            TextPointer? next = walker.GetLineStartPosition(1);
            if (next == null)
                break;

            lineIndex++;
            walker = next;
        }

        int columnOffset = lineStart.GetOffsetToPosition(position);
        return (lineIndex, Math.Max(0, columnOffset));
    }

    private static TextPointer RestorePosition(RichTextBox editor, int lineIndex, int columnOffset)
    {
        TextPointer contentStart = editor.Document.ContentStart;
        TextPointer? lineStart = contentStart.GetLineStartPosition(lineIndex, out int linesMoved);

        if (lineStart == null || linesMoved != lineIndex)
            lineStart = editor.Document.ContentEnd;

        TextPointer targetPosition = lineStart;

        if (columnOffset > 0)
        {
            TextPointer? columnPosition = lineStart.GetPositionAtOffset(columnOffset);
            if (columnPosition != null)
            {
                targetPosition = columnPosition;
            }
            else
            {
                TextPointer? lineEnd = lineStart.GetLineStartPosition(1) ?? editor.Document.ContentEnd;
                targetPosition = lineEnd;
            }
        }

        return targetPosition;
    }

    private sealed record SelectionSnapshot(
        (int lineIndex, int columnOffset) Start,
        (int lineIndex, int columnOffset) End,
        (int lineIndex, int columnOffset) Caret,
        bool IsEmpty);

    private static SelectionSnapshot CaptureSelection(RichTextBox editor)
    {
        TextSelection selection = editor.Selection;

        return new SelectionSnapshot(
            CapturePosition(editor, selection.Start),
            CapturePosition(editor, selection.End),
            CapturePosition(editor, editor.CaretPosition),
            selection.IsEmpty);
    }

    private static void RestoreSelection(RichTextBox editor, SelectionSnapshot snapshot)
    {
        TextPointer start = RestorePosition(editor, snapshot.Start.lineIndex, snapshot.Start.columnOffset);
        TextPointer end = RestorePosition(editor, snapshot.End.lineIndex, snapshot.End.columnOffset);
        TextPointer caret = RestorePosition(editor, snapshot.Caret.lineIndex, snapshot.Caret.columnOffset);

        if (snapshot.IsEmpty)
        {
            editor.CaretPosition = caret;
            return;
        }

        editor.Selection.Select(start, end);
        editor.CaretPosition = caret;
    }
}

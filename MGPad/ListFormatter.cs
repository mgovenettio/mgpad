using System.Collections.Generic;
using System.Linq;
using System.Windows.Documents;
using WpfRichTextBox = System.Windows.Controls.RichTextBox;

namespace MGPad;

internal static class ListFormatter
{

    private sealed record LineInfo(TextPointer Start, TextPointer End, string Text)
    {
        public string NormalizedText => Text.TrimEnd('\r', '\n');
    }

    public static void RenumberLists(WpfRichTextBox? editor)
    {
        if (editor?.Document is null)
            return;

        FlowDocument document = editor.Document;
        List<LineInfo> lines = GetLines(document).ToList();

        if (lines.Count == 0)
            return;

        SelectionSnapshot selectionSnapshot = CaptureSelection(editor);

        Dictionary<int, (ListLineType type, bool isUppercase, int count)> numberingState = new();

        for (int i = 0; i < lines.Count; i++)
        {
            if (!ListParser.TryParseListLine(lines[i].NormalizedText, out ParsedListLine parsedLine) ||
                (parsedLine.Type != ListLineType.Numbered && parsedLine.Type != ListLineType.Lettered))
            {
                continue;
            }

            // Drop numbering state for deeper indentation levels when moving back up the document tree.
            foreach (int level in numberingState.Keys.Where(level => level > parsedLine.IndentLevel).ToList())
            {
                numberingState.Remove(level);
            }

            int blockStart = i;
            List<ParsedListLine> blockLines = new() { parsedLine };

            i++;
            while (i < lines.Count &&
                   ListParser.TryParseListLine(lines[i].NormalizedText, out ParsedListLine nextLine) &&
                   nextLine.Type == parsedLine.Type &&
                   nextLine.IndentLevel == parsedLine.IndentLevel)
            {
                blockLines.Add(nextLine);
                i++;
            }

            int startingIndex = numberingState.TryGetValue(parsedLine.IndentLevel, out var state) &&
                                 state.type == parsedLine.Type &&
                                 state.isUppercase == parsedLine.IsUppercaseLetter
                ? state.count
                : 0;

            RenumberBlock(lines, blockStart, blockLines, parsedLine.Type, startingIndex);

            numberingState[parsedLine.IndentLevel] = (
                parsedLine.Type,
                parsedLine.IsUppercaseLetter,
                startingIndex + blockLines.Count);

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

    private static void RenumberBlock(
        IReadOnlyList<LineInfo> lines,
        int blockStart,
        IReadOnlyList<ParsedListLine> parsedLines,
        ListLineType listType,
        int startingIndex)
    {
        bool useUppercaseLetters = listType == ListLineType.Lettered && parsedLines[0].IsUppercaseLetter;
        char letterBase = useUppercaseLetters ? 'A' : 'a';

        for (int i = 0; i < parsedLines.Count; i++)
        {
            ParsedListLine parsedLine = parsedLines[i];
            LineInfo line = lines[blockStart + i];

            string indent = parsedLine.IndentText;
            string punctuation = parsedLine.Punctuation;
            string spacing = parsedLine.Spacing;

            string newMarker = listType switch
            {
                ListLineType.Numbered => (startingIndex + i + 1).ToString(),
                _ => BuildLetterMarker(letterBase, startingIndex + i)
            };

            string newPrefix = ListParser.BuildPrefix(indent, newMarker, punctuation, spacing);
            string existingPrefix = ListParser.BuildPrefix(indent, parsedLine.Marker, punctuation, spacing);

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

    private static (int lineIndex, int columnOffset) CapturePosition(WpfRichTextBox editor, TextPointer position)
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

    private static TextPointer RestorePosition(WpfRichTextBox editor, int lineIndex, int columnOffset)
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

    private static SelectionSnapshot CaptureSelection(WpfRichTextBox editor)
    {
        TextSelection selection = editor.Selection;

        return new SelectionSnapshot(
            CapturePosition(editor, selection.Start),
            CapturePosition(editor, selection.End),
            CapturePosition(editor, editor.CaretPosition),
            selection.IsEmpty);
    }

    private static void RestoreSelection(WpfRichTextBox editor, SelectionSnapshot snapshot)
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

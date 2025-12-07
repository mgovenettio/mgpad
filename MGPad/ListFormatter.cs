using System.Collections.Generic;
using System.Linq;
using System.Windows.Controls;
using System.Windows.Documents;

namespace MGPad;

internal static class ListFormatter
{

    private sealed record LineInfo(TextPointer Start, TextPointer End, string Text)
    {
        public string NormalizedText => Text.TrimEnd('\r', '\n');
    }

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
            if (!ListParser.TryMatchOrderedList(lines[i].NormalizedText, out OrderedListMatch firstMatch))
                continue;

            int blockStart = i;
            List<OrderedListMatch> blockMatches = new() { firstMatch };

            i++;
            while (i < lines.Count &&
                   ListParser.TryMatchOrderedList(lines[i].NormalizedText, out OrderedListMatch nextMatch) &&
                   nextMatch.Type == firstMatch.Type)
            {
                blockMatches.Add(nextMatch);
                i++;
            }

            RenumberBlock(lines, blockStart, blockMatches, firstMatch.Type);
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
        IReadOnlyList<OrderedListMatch> matches,
        ListLineType listType)
    {
        bool useUppercaseLetters = listType == ListLineType.Lettered && matches[0].IsUppercaseLetter;
        char letterBase = useUppercaseLetters ? 'A' : 'a';

        for (int i = 0; i < matches.Count; i++)
        {
            OrderedListMatch match = matches[i];
            LineInfo line = lines[blockStart + i];

            string indent = match.Indent;
            string punctuation = match.Punctuation;
            string spacing = match.Spacing;

            string newMarker = listType switch
            {
                ListLineType.Numbered => (i + 1).ToString(),
                _ => BuildLetterMarker(letterBase, i)
            };

            string newPrefix = ListParser.BuildPrefix(indent, newMarker, punctuation, spacing);
            string existingPrefix = match.ExistingPrefix;

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

using System.Text.RegularExpressions;

namespace MGPad;

internal enum ListLineType
{
    Numbered,
    Lettered,
    Bullet
}

internal sealed record ParsedListLine(
    string IndentText,
    int IndentLevel,
    ListLineType Type,
    string Marker,
    string Punctuation,
    string Spacing,
    string Content,
    string LineBreak,
    bool IsUppercaseLetter);

internal sealed record OrderedListMatch(
    ListLineType Type,
    Match Match,
    string ExistingPrefix,
    string Indent,
    string Punctuation,
    string Spacing,
    bool IsUppercaseLetter);

internal static class ListParser
{
    public const int IndentSpacesPerLevel = 4;

    private static readonly Regex NumberedListRegex = new(
        "^(?<prefix>(?<indent>\\s*)(?<marker>\\d+)(?<punct>[.)])(?<spacing>\\s+))",
        RegexOptions.Compiled);

    private static readonly Regex LetteredListRegex = new(
        "^(?<prefix>(?<indent>\\s*)(?<marker>[A-Za-z])(?<punct>[.)])(?<spacing>\\s+))",
        RegexOptions.Compiled);

    private static readonly Regex ListPrefixRegex = new(
        "^(?<indent>\\s*)(?:(?<number>\\d+)|(?<letter>[A-Za-z])|(?<bullet>[\\*-]))(?<punct>[.)])?\\s+",
        RegexOptions.Compiled);

    public static bool TryMatchOrderedList(string text, out OrderedListMatch matchInfo)
    {
        string normalizedLine = NormalizeLine(text);

        Match match = NumberedListRegex.Match(normalizedLine);
        if (match.Success)
        {
            matchInfo = CreateOrderedListMatch(ListLineType.Numbered, match, isUppercaseLetter: false);
            return true;
        }

        match = LetteredListRegex.Match(normalizedLine);
        if (match.Success)
        {
            bool isUppercase = char.IsUpper(match.Groups["marker"].Value[0]);
            matchInfo = CreateOrderedListMatch(ListLineType.Lettered, match, isUppercase);
            return true;
        }

        matchInfo = default!;
        return false;
    }

    public static bool TryParseListLine(string lineText, out ParsedListLine parsedLine)
    {
        string normalizedLine = NormalizeLine(lineText);
        Match match = ListPrefixRegex.Match(normalizedLine);

        if (!match.Success)
        {
            parsedLine = default!;
            return false;
        }

        string indentation = match.Groups["indent"].Value;
        string marker = match.Groups["number"].Success
            ? match.Groups["number"].Value
            : (match.Groups["letter"].Success
                ? match.Groups["letter"].Value
                : match.Groups["bullet"].Value);

        ListLineType type = match.Groups["number"].Success
            ? ListLineType.Numbered
            : (match.Groups["letter"].Success ? ListLineType.Lettered : ListLineType.Bullet);

        parsedLine = new ParsedListLine(
            indentation,
            GetIndentLevelFromIndent(indentation),
            type,
            marker,
            match.Groups["punct"].Value,
            match.Groups["spacing"].Value,
            normalizedLine[match.Length..],
            lineText[normalizedLine.Length..],
            match.Groups["letter"].Success && char.IsUpper(match.Groups["letter"].Value[0]));

        return true;
    }

    public static bool TryMatchListPrefix(string text, out Match match)
    {
        string normalizedLine = NormalizeLine(text);
        match = ListPrefixRegex.Match(normalizedLine);
        return match.Success;
    }

    public static bool IsListLine(string text)
    {
        return ListPrefixRegex.IsMatch(NormalizeLine(text));
    }

    public static int GetIndentLevel(string text)
    {
        return TryParseListLine(text, out ParsedListLine parsedLine)
            ? parsedLine.IndentLevel
            : 0;
    }

    public static int GetIndentLevelFromIndent(string indentation)
    {
        int spaces = 0;

        foreach (char character in indentation)
            spaces += character == '\t' ? IndentSpacesPerLevel : 1;

        return spaces / IndentSpacesPerLevel;
    }

    public static string SetIndentLevel(string text, int indentLevel)
    {
        string normalizedLine = NormalizeLine(text);
        string lineBreak = text[normalizedLine.Length..];

        Match match = ListPrefixRegex.Match(normalizedLine);
        if (!match.Success)
            return text;

        string updatedIndentation = new(' ', indentLevel * IndentSpacesPerLevel);
        string remainder = normalizedLine[match.Groups["indent"].Length..];

        return updatedIndentation + remainder + lineBreak;
    }

    public static string BuildLetterMarker(int index, bool useUppercase)
    {
        char letterBase = useUppercase ? 'A' : 'a';
        int zeroBasedIndex = Math.Max(0, index - 1);
        int maxOffset = (useUppercase ? 'Z' : 'z') - letterBase;
        int clamped = Math.Min(zeroBasedIndex, maxOffset);

        return ((char)(letterBase + clamped)).ToString();
    }

    public static string BuildPrefix(string indent, string marker, string punctuation, string spacing)
    {
        return indent + marker + punctuation + spacing;
    }

    public static string NormalizePrefix(Match match, string marker)
    {
        return BuildPrefix(
            match.Groups["indent"].Value,
            marker,
            match.Groups["punct"].Value,
            match.Groups["spacing"].Value);
    }

    public static string NormalizeLine(string text)
    {
        return text.TrimEnd('\r', '\n');
    }

    private static OrderedListMatch CreateOrderedListMatch(ListLineType type, Match match, bool isUppercaseLetter)
    {
        return new OrderedListMatch(
            type,
            match,
            match.Groups["prefix"].Value,
            match.Groups["indent"].Value,
            match.Groups["punct"].Value,
            match.Groups["spacing"].Value,
            isUppercaseLetter);
    }
}

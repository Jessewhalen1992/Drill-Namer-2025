using System.Text.RegularExpressions;

namespace DrillNamer.Core;

/// <summary>
/// Helper methods for working with attribute text.
/// </summary>
public static class AttributeUtils
{
    /// <summary>
    /// Normalizes text by removing whitespace and converting to upper case.
    /// </summary>
    public static string NormalizeText(string text)
    {
        text = text ?? string.Empty;
        text = Regex.Replace(text, "\s+", string.Empty);
        return text.ToUpperInvariant();
    }

    /// <summary>
    /// Determines whether the provided text matches a drill tag pattern.
    /// Pattern: d-d-ddd-d.
    /// </summary>
    public static bool MatchesDrillTag(string text)
    {
        if (string.IsNullOrEmpty(text))
            return false;
        return Regex.IsMatch(text, @"^\d{1,2}-\d{1,2}-\d{1,3}-\d{1,2}$");
    }
}

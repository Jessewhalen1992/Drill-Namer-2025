using DrillNamer.Core;
using NUnit.Framework;

namespace DrillNamer.Tests;

public class AttributeUtilsTests
{
    [TestCase(" a b c ", "ABC")]
    [TestCase("A-B-C", "A-B-C")]
    public void NormalizeText_RemovesWhitespaceAndUppercases(string input, string expected)
    {
        Assert.AreEqual(expected, AttributeUtils.NormalizeText(input));
    }

    [TestCase("1-2-3-4", true)]
    [TestCase("01-02-003-04", true)]
    [TestCase("1-2-3", false)]
    public void MatchesDrillTag_DetectsPattern(string input, bool expected)
    {
        Assert.AreEqual(expected, AttributeUtils.MatchesDrillTag(input));
    }
}

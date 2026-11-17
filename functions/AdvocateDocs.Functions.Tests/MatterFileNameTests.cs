using AdvocateDocs.Functions.MatterFolders;

namespace AdvocateDocs.Functions.Tests;

public class MatterFileNameTests
{
    private static readonly DateTimeOffset Received = new(2026, 11, 13, 23, 30, 0, TimeSpan.FromHours(-5));

    [Fact]
    public void For_Should_PrefixUtcDate_When_ReceivedWithOffset()
        => Assert.Equal("2026-11-14 Hello.eml", MatterFileName.For(Received, "Hello"));

    [Fact]
    public void For_Should_ReplaceIllegalChars_When_SubjectHasThem()
        => Assert.Equal("2026-11-14 Re- a-b-c-d-e-f-g-h-i.eml", MatterFileName.For(Received, "Re: a\"b*c<d>e?f/g\\h|i"));

    [Fact]
    public void For_Should_CollapseWhitespaceAndTrimDots_When_SubjectIsMessy()
        => Assert.Equal("2026-11-14 a b.eml", MatterFileName.For(Received, " ..a   b.. "));

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("  . ")]
    public void For_Should_UseNoSubject_When_SubjectIsEmpty(string? subject)
        => Assert.Equal("2026-11-14 (no subject).eml", MatterFileName.For(Received, subject));

    [Fact]
    public void For_Should_CapAt250Chars_When_SubjectIsLong()
    {
        var name = MatterFileName.For(Received, new string('a', 400) + " tail");
        Assert.Equal(250, name.Length);
        Assert.EndsWith("a.eml", name);
    }

    [Fact]
    public void For_Should_TrimEndAfterCut_When_CutLandsOnSpace()
    {
        var name = MatterFileName.For(Received, new string('a', 235) + " b");
        Assert.Equal(11 + 235 + 4, name.Length);
    }
}

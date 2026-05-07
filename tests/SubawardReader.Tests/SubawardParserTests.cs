using Xunit;

namespace SubawardReader.Tests;

public class SubawardParserTests
{
    [Fact]
    public void Should_Contain_Expected_Subrecipients()
    {
        var expectedNames = new List<string>
        {
            "Indiana",
            "Mayo",
            "Purdue",
            "Florida"
        };

        Assert.Contains("Indiana", expectedNames);
        Assert.Contains("Mayo", expectedNames);
        Assert.Contains("Purdue", expectedNames);
        Assert.Contains("Florida", expectedNames);
    }
}
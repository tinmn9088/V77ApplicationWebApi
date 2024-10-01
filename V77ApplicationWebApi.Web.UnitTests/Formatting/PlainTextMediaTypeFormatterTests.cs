using Xunit;

namespace V77ApplicationWebApi.Web.Formatting.UnitTests;

public class PlainTextMediaTypeFormatterTests
{
    private readonly PlainTextMediaTypeFormatter formatter;

    public PlainTextMediaTypeFormatterTests() => this.formatter = new();

    [Fact]
    public void PlainTextMediaTypeFormatter_ShouldSetSupportedMediaType()
    {
        Assert.Single(this.formatter.SupportedMediaTypes);
        Assert.Contains(PlainTextMediaTypeFormatter.DefaultMediaType, this.formatter.SupportedMediaTypes);
    }
}

using System;
using System.IO;
using System.Threading.Tasks;
using V77ApplicationWebApi.Web.UnitTests.Common;
using Xunit;

namespace V77ApplicationWebApi.Web.Formatting.UnitTests;

public class PlainTextMediaTypeFormatterTests
{
    private readonly PlainTextMediaTypeFormatter _formatter;

    public PlainTextMediaTypeFormatterTests() => _formatter = new();

    [Fact]
    public void PlainTextMediaTypeFormatter_ShouldSetSupportedMediaType()
    {
        _ = Assert.Single(_formatter.SupportedMediaTypes);
        Assert.Contains(PlainTextMediaTypeFormatter.DefaultMediaType, _formatter.SupportedMediaTypes);
    }

    [Theory]
    [InlineData(typeof(int))]
    [InlineData(typeof(string))]
    [InlineData(typeof(object))]
    public void CanReadType_ReturnsFalseTheory(Type type) => Assert.False(_formatter.CanReadType(type));

    [Theory]
    [InlineData(typeof(char))]
    [InlineData(typeof(char[]))]
    [InlineData(typeof(string))]
    public void CanWriteType_WhenSupported_ReturnsTrueTheory(Type type) => Assert.True(_formatter.CanWriteType(type));

    [Theory]
    [InlineData(typeof(int))]
    [InlineData(typeof(object))]
    public void CanWriteType_WhenNotSupported_ReturnsFalseTheory(Type type) => Assert.False(_formatter.CanWriteType(type));

    [Fact]
    public async Task WriteToStreamAsync_WhenValueIsString_ShouldWriteToStream()
    {
        // Arrange
        string value = "TestString";

        // Act
        string read = await WriteAndReadStreamAsync(value);

        // Assert
        Assert.Equal(value, read);
    }

    [Fact]
    public async Task WriteToStreamAsync_WhenValueIsChar_ShouldWriteToStream()
    {
        // Arrange
        char value = 'A';

        // Act
        string read = await WriteAndReadStreamAsync(value);

        // Assert
        Assert.Equal(value.ToString(), read);
    }

    [Fact]
    public async Task WriteToStreamAsync_WhenValueIsCharArray_ShouldWriteToStream()
    {
        // Arrange
        char[] value = ['A', 'B', 'C'];

        // Act
        string read = await WriteAndReadStreamAsync(value);

        // Assert
        Assert.Equal(new string(value), read);
    }

    [Fact]
    public async Task WriteToStreamAsync_WhenValueIsNotSupported_ThrowsNotSupportedException()
    {
        // Arrange
        int valueOfUnsupportedType = 42;

        // Act, Assert
        _ = await Assert.ThrowsAsync<NotSupportedException>(() => WriteAndReadStreamAsync(valueOfUnsupportedType));
    }

    private async Task<string> WriteAndReadStreamAsync(object value)
    {
        using IgnoreCloseMemoryStream writeStream = new(timesToIgnoreClose: 2);

        await _formatter.WriteToStreamAsync(type: default, value, writeStream, content: default, transportContext: default);

        writeStream.Position = 0;
        using StreamReader reader = new(writeStream);

        return await reader.ReadToEndAsync();
    }
}

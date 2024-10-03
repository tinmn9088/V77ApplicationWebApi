using System;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Net.Http.Formatting;
using System.Net.Http.Headers;
using System.Threading.Tasks;

namespace V77ApplicationWebApi.Web.Formatting;

/// <summary>
/// Represents the <see cref="MediaTypeFormatter"/> class to handle plain text.
/// </summary>
public class PlainTextMediaTypeFormatter : MediaTypeFormatter
{
    /// <summary>
    /// Gets the default media type for plain text, namely "text/plain".
    /// </summary>
    public static readonly MediaTypeHeaderValue DefaultMediaType = new("text/plain");

    public PlainTextMediaTypeFormatter() => SupportedMediaTypes.Add(DefaultMediaType);

    public override bool CanReadType(Type type) => false;

    public override bool CanWriteType(Type type) => type == typeof(string) || type == typeof(char) || type == typeof(char[]);

    /// <summary>
    /// Called during serialization to write an object of the specified type to the specified stream.
    /// </summary>
    /// <param name="type">The type of the object to write.</param>
    /// <param name="value">The object to write.</param>
    /// <param name="writeStream">The stream to write to.</param>
    /// <param name="content">The <see cref="HttpContent"/> for the content being written.</param>
    /// <param name="transportContext">The transport context.</param>
    /// <returns>A <see cref="Task"/> representing the asynchronous operation.</returns>
    /// <exception cref="NotSupportedException">
    /// When <paramref name="value"/> type is neither <see cref="string"/>,
    /// nor <see cref="char"/>, nor <see cref="char[]"/>.
    /// </exception>
    public override async Task WriteToStreamAsync(Type type, object value, Stream writeStream, HttpContent content, TransportContext transportContext)
    {
        using StreamWriter streamWriter = new(writeStream);

        Task writeTask = value switch
        {
            string valueString => streamWriter.WriteAsync(valueString),
            char valueChar => streamWriter.WriteAsync(valueChar),
            char[] valueCharArray => streamWriter.WriteAsync(valueCharArray),
            _ => throw new NotSupportedException(),
        };

        await writeTask.ConfigureAwait(false);
    }
}

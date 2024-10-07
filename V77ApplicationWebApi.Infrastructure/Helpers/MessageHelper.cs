using System.Text;

namespace V77ApplicationWebApi.Infrastructure.Helpers;

internal static class MessageHelper
{
    private const string NullString = "null";

    public static string BuildArgsString(object?[]? args)
    {
        if (args is null)
        {
            return NullString;
        }

        StringBuilder result = new();

        for (int i = 0; i < args.Length; i++)
        {
            if (i > 0)
            {
                _ = result.Append(' ');
            }

            _ = result
                .Append('{')
                .Append(args[i]?.ToString() ?? NullString)
                .Append('}');
        }

        return result.ToString();
    }
}

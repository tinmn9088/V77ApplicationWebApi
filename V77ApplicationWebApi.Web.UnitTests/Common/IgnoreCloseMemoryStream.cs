using System.IO;

namespace V77ApplicationWebApi.Web.UnitTests.Common;

/// <summary>
/// <see cref="StreamWriter"/> and <see cref="StreamReader"/> usually close
/// used <see cref="Stream"/>. This class is used as a workaround
/// to leave the <see cref="Stream"/> open despite <see cref="Stream.Close"/>
/// is called.
/// </summary>
internal class IgnoreCloseMemoryStream(int timesToIgnoreClose)
    : MemoryStream
{
    public int TimesToIgnoreClose { get; private set; } = timesToIgnoreClose;

    public override void Close()
    {
        if (TimesToIgnoreClose > 0)
        {
            TimesToIgnoreClose--;
        }
        else
        {
            base.Close();
        }
    }
}

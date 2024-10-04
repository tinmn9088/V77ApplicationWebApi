using System.Threading;
using System.Threading.Tasks;

namespace V77ApplicationWebApi.Core;

public interface IV77ApplicationConnectionFactory
{
    ValueTask<IV77ApplicationConnection> GetConnectionAsync(ConnectionProperties properties, CancellationToken cancellationToken);
}

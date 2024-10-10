using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Extensions.Logging;
using V77ApplicationWebApi.Core;

namespace V77ApplicationWebApi.Infrastructure;

public class ComV77ApplicationConnectionFactory(
    IInstanceFactory instanceFactory,
    IMemberInvoker memberInvoker,
    ILogger<ComV77ApplicationConnection> connectionLogger)
    : IV77ApplicationConnectionFactory
{
    private readonly IDictionary<ConnectionProperties, ComV77ApplicationConnection> _connections = new Dictionary<ConnectionProperties, ComV77ApplicationConnection>();

    private readonly SemaphoreSlim _factoryLock = new(1, 1);

    public int InstancesCount => _connections.Count;

    private Func<ConnectionProperties, ValueTask> LockFactoryAndRemoveConnection => async (ConnectionProperties properties) =>
    {
        await _factoryLock.WaitAsync().ConfigureAwait(false);
        _ = _connections.Remove(properties);
    };

    private Action UnlockFactory => () => _ = _factoryLock.Release();

    public async ValueTask<IV77ApplicationConnection> GetConnectionAsync(ConnectionProperties properties, CancellationToken cancellationToken)
    {
        await _factoryLock.WaitAsync(cancellationToken).ConfigureAwait(false);

        try
        {
            if (!_connections.TryGetValue(properties, out ComV77ApplicationConnection connection))
            {
                connection = new(
                    properties,
                    instanceFactory,
                    memberInvoker,
                    logger: connectionLogger,
                    beforeDisposeCallback: LockFactoryAndRemoveConnection,
                    afterDisposeCallback: UnlockFactory);

                _connections.Add(properties, connection);
            }

            return connection;
        }
        finally
        {
            _ = _factoryLock.Release();
        }
    }
}

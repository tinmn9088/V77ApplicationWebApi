using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using V77ApplicationWebApi.Core;

namespace V77ApplicationWebApi.Infrastructure;

public class V77ApplicationConnection : IV77ApplicationConnection
{
    public ValueTask ConnectAsync(CancellationToken cancellationToken) => throw new System.NotImplementedException();

    public ValueTask DisposeAsync() => throw new System.NotImplementedException();

    public ValueTask RunErtAsync(string ertRelativePath, CancellationToken cancellationToken) => throw new System.NotImplementedException();

    public ValueTask RunErtAsync(string ertRelativePath, IReadOnlyDictionary<string, string> ertContext, CancellationToken cancellationToken) => throw new System.NotImplementedException();

    public ValueTask<string> RunErtAsync(string ertRelativePath, IReadOnlyDictionary<string, string> ertContext, string resultName, CancellationToken cancellationToken) => throw new System.NotImplementedException();

    public ValueTask<string> RunErtAsync(string ertRelativePath, IReadOnlyDictionary<string, string> ertContext, string resultName, string errorMessageName, CancellationToken cancellationToken) => throw new System.NotImplementedException();
}

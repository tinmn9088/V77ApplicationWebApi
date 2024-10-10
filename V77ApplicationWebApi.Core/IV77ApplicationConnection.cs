using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using V77ApplicationWebApi.Core.Exceptions;

namespace V77ApplicationWebApi.Core;

public interface IV77ApplicationConnection : IAsyncDisposable
{
    ConnectionProperties Properties { get; }

    /// <summary>
    /// Connect to infobase.
    /// </summary>
    /// <param name="cancellationToken">The token to cancel the operation.</param>
    /// <returns>A <see cref="ValueTask"/> representing the asynchronous operation.</returns>
    /// <exception cref="FailedToConnectException">If failed to connect to infobase.</exception>
    /// <exception cref="ErrorsCountExceededException">If too many errors has with connection.</exception>
    /// <exception cref="OperationCanceledException">If <paramref name="cancellationToken"/> was cancelled.</exception>
    /// <exception cref="ObjectDisposedException">Connection is already disposed.</exception>
    ValueTask ConnectAsync(CancellationToken cancellationToken);

    /// <summary>
    /// Run ERT specified by <paramref name="ertRelativePath"/>.
    /// </summary>
    /// <param name="ertRelativePath">ERT path relative to infobase path.</param>
    /// <param name="cancellationToken">The token to cancel the operation.</param>
    /// <returns>A <see cref="ValueTask"/> representing the asynchronous operation.</returns>
    /// <exception cref="FailedToRunErtException">If an error has occurred while running ERT.</exception>
    /// <exception cref="ErrorsCountExceededException">If too many errors has with connection.</exception>
    /// <exception cref="OperationCanceledException">If <paramref name="cancellationToken"/> was cancelled.</exception>
    /// <exception cref="ObjectDisposedException">Connection is already disposed.</exception>
    ValueTask RunErtAsync(string ertRelativePath, CancellationToken cancellationToken);

    /// <summary>
    /// Run ERT specified by <paramref name="ertRelativePath"/> with parameters from <paramref name="ertContext"/>.
    /// </summary>
    /// <param name="ertContext">Parameters that will be passed to ERT.</param>
    /// <inheritdoc cref="RunErtAsync(string, CancellationToken)"/>
    ValueTask RunErtAsync(string ertRelativePath, IReadOnlyDictionary<string, string>? ertContext, CancellationToken cancellationToken);

    /// <summary>
    /// Run ERT specified by <paramref name="ertRelativePath"/> with parameters from <paramref name="ertContext"/>
    /// and get result by <paramref name="resultName"/>.
    /// </summary>
    /// <param name="resultName">Name by which the result of ERT run can be retrieved.</param>
    /// <inheritdoc cref="RunErtAsync(string, IReadOnlyDictionary{string, string}?, CancellationToken)"/>
    ValueTask<string?> RunErtAsync(string ertRelativePath, IReadOnlyDictionary<string, string>? ertContext, string? resultName, CancellationToken cancellationToken);

    /// <param name="errorMessageName">Name by which the error message of ERT run can be retrieved.</param>
    /// <exception cref="ErtReturnedErrorMessageException">
    /// Value retrieved by <paramref name="errorMessageName"/> after run was not <c>null</c>.
    /// </exception>
    /// <inheritdoc cref="RunErtAsync(string, IReadOnlyDictionary{string, string}?, string?, CancellationToken)"/>
    ValueTask<string?> RunErtAsync(string ertRelativePath, IReadOnlyDictionary<string, string>? ertContext, string? resultName, string? errorMessageName, CancellationToken cancellationToken);
}

using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Extensions.Logging;
using V77ApplicationWebApi.Core;
using V77ApplicationWebApi.Infrastructure.Exceptions;
using static V77ApplicationWebApi.Infrastructure.Logging.LoggerExtensions;

namespace V77ApplicationWebApi.Infrastructure;

public class ComV77ApplicationConnection : IV77ApplicationConnection
{
    private readonly IInstanceFactory _instanceFactory;

    private readonly IMemberInvoker _memberInvoker;

    private readonly ILogger<ComV77ApplicationConnection> _logger;

    private readonly SemaphoreSlim _connectionLock;

    private Type? _comObjectType;

    private object? _comObject;

    private bool _isInitialized;

    private int _errorsCount;

    internal ComV77ApplicationConnection(
        ConnectionProperties properties,
        IInstanceFactory instanceFactory,
        IMemberInvoker memberInvoker,
        ILogger<ComV77ApplicationConnection> logger)
    {
        Properties = properties;

        _memberInvoker = memberInvoker;
        _instanceFactory = instanceFactory;
        _logger = logger;

        _connectionLock = new(1);
        _isInitialized = false;
        _errorsCount = 0;
    }

    public static TimeSpan InitializeTimeout => TimeSpan.FromSeconds(30);

    public static string ComObjectTypeName => "V77.Application";

    public ConnectionProperties Properties { get; }

    private static int MaxErrorsCount => 3;

    public async ValueTask ConnectAsync(CancellationToken cancellationToken)
    {
        _logger.LogTryingConnect(Properties.InfobasePath);

        await _connectionLock.WaitAsync(cancellationToken).ConfigureAwait(false);

        try
        {
            CheckState(isInitializing: true);

            _comObjectType ??= GetComObjectType();
            _comObject ??= CreateComObject();

            if (!_isInitialized)
            {
                _isInitialized = await InitializeAsync(cancellationToken).ConfigureAwait(false);
            }
        }
        catch (Exception)
        {
            _errorsCount++;
            throw;
        }
        finally
        {
            _ = _connectionLock.Release();
        }
    }

    public ValueTask DisposeAsync() => throw new NotImplementedException();

    public ValueTask RunErtAsync(string ertRelativePath, CancellationToken cancellationToken) => throw new NotImplementedException();

    public ValueTask RunErtAsync(string ertRelativePath, IReadOnlyDictionary<string, string> ertContext, CancellationToken cancellationToken) => throw new NotImplementedException();

    public ValueTask<string> RunErtAsync(string ertRelativePath, IReadOnlyDictionary<string, string> ertContext, string resultName, CancellationToken cancellationToken) => throw new NotImplementedException();

    public ValueTask<string> RunErtAsync(string ertRelativePath, IReadOnlyDictionary<string, string> ertContext, string resultName, string errorMessageName, CancellationToken cancellationToken) => throw new NotImplementedException();

    /// <summary>
    /// Confirm that <see cref="_errorsCount"/> is not exceeded and COM-object
    /// is created and initialized.
    /// </summary>
    /// <param name="isInitializing">When <c>true</c> only <see cref="_errorsCount"/> value is checked.</param>
    /// <exception cref="ErrorsCountExceededException">If <see cref="_errorsCount"/> is exceeded.</exception>
    /// <exception cref="InvalidOperationException">
    /// If <see cref="_comObject"/> is not created yet, or <see cref="_isInitialized"/> is <c>false</c>.
    /// </exception>
    private void CheckState(bool isInitializing = false)
    {
        if (_errorsCount >= MaxErrorsCount)
        {
            throw new ErrorsCountExceededException(_errorsCount);
        }

        if (!isInitializing)
        {
            if (_comObject is null)
            {
                throw new InvalidOperationException("COM object is not created");
            }

            if (!_isInitialized)
            {
                throw new InvalidOperationException("COM object is not connected to infobase");
            }
        }
    }

    /// <summary>
    /// Run in parallel:
    /// <list type="number">
    /// <item>Initializing <see cref="_comObject"/></item>
    /// <item>Timer to detect if <see cref="InitializeTimeout"/> is exceeded</item>
    /// </list>
    /// </summary>
    /// <param name="cancellationToken">The token to monitor for cancellation.</param>
    /// <returns><c>true</c> if initialization was successful.</returns>
    /// <exception cref="InitializeTimeoutExceededException">
    /// If initialization took more time than <see cref="InitializeTimeout"/>.
    /// </exception>
    private async ValueTask<bool> InitializeAsync(CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();

        // Lock that guarantees that the timeoutExceededTask starts after initializeTask
        SemaphoreSlim initializeTaskStarted = new(0, 1);

        // Initialization task
        Task<bool> initializeTask = Task.Run(
            () =>
            {
                // Let the timeoutExceededTask begin
                _ = initializeTaskStarted.Release();

                cancellationToken.ThrowIfCancellationRequested();

                // Get "RMTrade" property value
                object? rmtrade = GetPropertyValue(
                    target: _comObject,
                    propertyName: "RMTrade",
                    isInitializing: true);

                cancellationToken.ThrowIfCancellationRequested();

                // Invoke "Initialize" method
                return (bool)InvokeMethod(
                    target: _comObject,
                    methodName: "Initialize",
                    args: [rmtrade!, $"/D{Properties.InfobasePath} /N{Properties.Username} /P{Properties.Password}", "NO_SPLASH_SHOW"],
                    isInitializing: true);
            }, cancellationToken);

        // Timer task
        Task timeoutExceededTask = Task.Run(
            async () =>
            {
                await initializeTaskStarted.WaitAsync(cancellationToken).ConfigureAwait(false);
                await Task.Delay(InitializeTimeout, cancellationToken).ConfigureAwait(false);
            }, cancellationToken);

        // Confirm that initializeTask completed before timeoutExceededTask
        bool timeoutExceeded = await Task.WhenAny(initializeTask, timeoutExceededTask).ConfigureAwait(false) != initializeTask;

        return timeoutExceeded
            ? throw new InitializeTimeoutExceededException(InitializeTimeout)
            : await initializeTask;
    }

    /// <inheritdoc cref="IInstanceFactory.GetTypeFromProgID(string)"/>
    private Type GetComObjectType() => _instanceFactory.GetTypeFromProgID(ComObjectTypeName);

    /// <inheritdoc cref="IInstanceFactory.CreateInstance(Type)"/>
    private object? CreateComObject() => _instanceFactory.CreateInstance(_comObjectType);

    /// <inheritdoc cref="CheckStateAndHandleError(Func{object?}, bool)"/>
    /// <inheritdoc cref="IMemberInvoker.InvokePublicMethodByName(object, string, object[]?)"/>
    private object? InvokeMethod(object target, string methodName, object[]? args, bool isInitializing = false) =>
        CheckStateAndHandleError(() => _memberInvoker.InvokePublicMethodByName(target, methodName, args), isInitializing);

    /// <inheritdoc cref="CheckStateAndHandleError(Func{object?}, bool)"/>
    /// <inheritdoc cref="IMemberInvoker.GetPropertyValueByName(object, string)"/>
    private object? GetPropertyValue(object target, string propertyName, bool isInitializing = false) =>
        CheckStateAndHandleError(() => _memberInvoker.GetPropertyValueByName(target, propertyName), isInitializing);

    /// <param name="isInitializing"><c>true</c> if called within <see cref="InitializeAsync(CancellationToken)"/>.</param>
    private object? CheckStateAndHandleError(Func<object?> getValue, bool isInitializing)
    {
        CheckState(isInitializing: isInitializing);

        try
        {
            return getValue();
        }
        catch (Exception)
        {
            if (!isInitializing)
            {
                _errorsCount++;
            }

            throw;
        }
    }
}

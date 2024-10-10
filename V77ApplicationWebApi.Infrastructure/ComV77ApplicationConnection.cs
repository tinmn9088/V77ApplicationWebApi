using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Extensions.Logging;
using V77ApplicationWebApi.Core;
using V77ApplicationWebApi.Core.Exceptions;
using V77ApplicationWebApi.Infrastructure.Exceptions;
using static V77ApplicationWebApi.Infrastructure.Logging.LoggerExtensions;

namespace V77ApplicationWebApi.Infrastructure;

public sealed class ComV77ApplicationConnection : IV77ApplicationConnection
{
    private readonly IInstanceFactory _instanceFactory;

    private readonly IMemberInvoker _memberInvoker;

    private readonly ILogger<ComV77ApplicationConnection> _logger;

    private readonly SemaphoreSlim _connectionLock;

    private readonly CancellationTokenSource _disposeTokenSource;

    private readonly SemaphoreSlim _disposeLock;

    private Type? _comObjectType;

    private object? _comObject;

    private bool _isInitialized;

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

        _connectionLock = new(1, 1);
        _isInitialized = false;

        _disposeTokenSource = new();
        _disposeLock = new(1, 1);

        ComObjectErrorsCount = 0;
    }

    public static TimeSpan DefaultInitializeTimeout => TimeSpan.FromSeconds(30);

    public static TimeSpan DefaultDisposeTimeout => TimeSpan.FromSeconds(5);

    public static string ComObjectTypeName => "V77.Application";

    public static int MaxComObjectErrorsCount => 3;

    public ConnectionProperties Properties { get; }

    public int ComObjectErrorsCount { get; private set; }

    public async ValueTask ConnectAsync(CancellationToken cancellationToken)
    {
        _logger.LogTryingConnect(infobasePath: Properties.InfobasePath);

        StopDisposeTimer();

        await _connectionLock.WaitAsync(cancellationToken).ConfigureAwait(false);

        try
        {
            CheckState(isInitializing: true);

            _comObjectType ??= GetComObjectType();
            _comObject ??= CreateComObject();

            cancellationToken.ThrowIfCancellationRequested();

            if (!_isInitialized)
            {
                _logger.LogInitializingConnection(infobasePath: Properties.InfobasePath);
                _isInitialized = await InitializeAsync(cancellationToken).ConfigureAwait(false);
            }
            else
            {
                _logger.LogAlreadyConnected(infobasePath: Properties.InfobasePath);
            }
        }
        catch (OperationCanceledException)
        {
            throw;
        }
        catch (ErrorsCountExceededException)
        {
            throw;
        }
        catch (Exception ex)
        {
            ComObjectErrorsCount++;
            throw new FailedToConnectException(Properties.InfobasePath, ex);
        }
        finally
        {
            _ = _connectionLock.Release();
        }
    }

    public async ValueTask DisposeAsync()
    {
        StartDisposeTimer(out TimeSpan disposeTimeout);

        // Allow only one thread to enter
        if (await _disposeLock.WaitAsync(TimeSpan.Zero).ConfigureAwait(false))
        {
            // Wait until dispose timer elapses
            await WaitDisposeTimerAsync();

            // Actual dispose
            await _connectionLock.WaitAsync().ConfigureAwait(false);

            _logger.LogConnectionDisposing(infobasePath: Properties.InfobasePath, disposeTimeout);

            ReleaseComObject();
            _disposeTokenSource.Dispose();
            _disposeLock.Dispose();
            _connectionLock.Dispose();

            _logger.LogConnectionDisposed(infobasePath: Properties.InfobasePath);
        }
    }

    public async ValueTask RunErtAsync(string ertRelativePath, CancellationToken cancellationToken) =>
        await RunErtAsync(ertRelativePath, ertContext: default, cancellationToken);

    public async ValueTask RunErtAsync(string ertRelativePath, IReadOnlyDictionary<string, string> ertContext, CancellationToken cancellationToken) =>
        _ = await RunErtAsync(ertRelativePath, ertContext, resultName: default, cancellationToken).ConfigureAwait(false);

    public ValueTask<string?> RunErtAsync(string ertRelativePath, IReadOnlyDictionary<string, string> ertContext, string resultName, CancellationToken cancellationToken) =>
        RunErtAsync(ertRelativePath, ertContext, resultName, errorMessageName: default, cancellationToken);

    public async ValueTask<string?> RunErtAsync(string ertRelativePath, IReadOnlyDictionary<string, string> ertContext, string resultName, string errorMessageName, CancellationToken cancellationToken)
    {
        _logger.LogRunningErt(infobasePath: Properties.InfobasePath, ertRelativePath);

        StopDisposeTimer();

        await _connectionLock.WaitAsync(cancellationToken).ConfigureAwait(false);

        try
        {
            string ertFullPath = Path.Combine(Properties.InfobasePath, ertRelativePath);

            // CreateObject("ValueList")
            object contextValueList = InvokeMethod(
                target: _comObject,
                methodName: "CreateObject",
                args: ["ValueList"]);

            if (ertContext is not null)
            {
                foreach (KeyValuePair<string, string> ertContextEntry in ertContext)
                {
                    cancellationToken.ThrowIfCancellationRequested();

                    // ValueList.AddValue(Value, Name)
                    _ = InvokeMethod(
                        target: contextValueList,
                        methodName: "AddValue",
                        args: [ertContextEntry.Value, ertContextEntry.Key]);
                }
            }

            cancellationToken.ThrowIfCancellationRequested();

            // OpenForm(ObjectName, Context, FullPath)
            _ = InvokeMethod(
                target: _comObject,
                methodName: "OpenForm",
                args: ["Report", contextValueList, ertFullPath]);

            if (errorMessageName is not null)
            {
                cancellationToken.ThrowIfCancellationRequested();

                // ValueList.Get(Name)
                object? errorMessage = InvokeMethod(
                    target: contextValueList,
                    methodName: "Get",
                    args: [errorMessageName]);

                if (errorMessage is not null)
                {
                    throw new ErtReturnedErrorMessageException(Properties.InfobasePath, errorMessage.ToString());
                }
            }

            if (resultName is not null)
            {
                cancellationToken.ThrowIfCancellationRequested();

                // ValueList.Get(Name)
                object? result = InvokeMethod(
                    target: contextValueList,
                    methodName: "Get",
                    args: [resultName]);

                return result?.ToString();
            }
            else
            {
                return default;
            }
        }
        catch (OperationCanceledException)
        {
            throw;
        }
        catch (ErrorsCountExceededException)
        {
            throw;
        }
        catch (Exception ex)
        {
            throw new FailedToRunErtException(Properties.InfobasePath, ex);
        }
        finally
        {
            _ = _connectionLock.Release();
        }
    }

    /// <summary>
    /// Confirm that <see cref="ComObjectErrorsCount"/> is not exceeded and COM-object is created and initialized.
    /// </summary>
    /// <param name="isInitializing">When <c>true</c> only <see cref="ComObjectErrorsCount"/> value is checked.</param>
    /// <exception cref="ErrorsCountExceededException">If <see cref="ComObjectErrorsCount"/> is exceeded.</exception>
    /// <exception cref="InvalidOperationException">
    /// If <see cref="_comObject"/> is not created yet, or <see cref="_isInitialized"/> is <c>false</c>.
    /// </exception>
    private void CheckState(bool isInitializing = false)
    {
        if (ComObjectErrorsCount >= MaxComObjectErrorsCount)
        {
            throw new ErrorsCountExceededException(ComObjectErrorsCount);
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
    /// <item>
    /// Initializing <see cref="_comObject"/>
    /// </item>
    /// <item>
    /// Timer to detect if <see cref="ConnectionProperties.InitializeTimeout"/>
    /// or <see cref="DefaultInitializeTimeout"/> is exceeded
    /// </item>
    /// </list>
    /// </summary>
    /// <param name="cancellationToken">The token to monitor for cancellation.</param>
    /// <returns><c>true</c> if initialization was successful.</returns>
    /// <exception cref="InitializeTimeoutExceededException">
    /// If initialization took more time than the limit.
    /// </exception>
    /// <exception cref="OperationCanceledException">If <paramref name="cancellationToken"/> was cancelled.</exception>
    /// <exception cref="FailedToInitializeException">If "Initialize" method returned false.</exception>
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
                bool isInitialized = (bool)InvokeMethod(
                    target: _comObject,
                    methodName: "Initialize",
                    args: [rmtrade, $"/D{Properties.InfobasePath} /N{Properties.Username} /P{Properties.Password}", "NO_SPLASH_SHOW"],
                    isInitializing: true);

                return isInitialized
                    ? isInitialized
                    : throw new FailedToInitializeException(Properties.InfobasePath);
            }, cancellationToken);

        // Prepare cancellation token for timer task
        using CancellationTokenSource timeoutExceededTaskCancellationTokenSource = new();
        CancellationToken timeoutExceededTaskCancellationToken = timeoutExceededTaskCancellationTokenSource.Token;

        // Timer task
        Task timeoutExceededTask = Task.Run(
            async () =>
            {
                await initializeTaskStarted.WaitAsync(cancellationToken).ConfigureAwait(false);
                await Task.Delay(Properties.InitializeTimeout ?? DefaultInitializeTimeout, timeoutExceededTaskCancellationToken).ConfigureAwait(false);
            }, timeoutExceededTaskCancellationToken);

        // Confirm that initialization task completed before timer task
        bool timeoutExceeded = await Task.WhenAny(initializeTask, timeoutExceededTask).ConfigureAwait(false) != initializeTask;

        // Explicitly cancel timer task
        timeoutExceededTaskCancellationTokenSource.Cancel();

        return timeoutExceeded
            ? throw new InitializeTimeoutExceededException(DefaultInitializeTimeout)
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

    /// <remarks>If <paramref name="isInitializing"/> is <c>false</c> increments <see cref="ComObjectErrorsCount"/> on errors
    /// (excluding <see cref="ErrorsCountExceededException"/> and <see cref="InvalidOperationException"/>).</remarks>
    /// <param name="isInitializing"><c>true</c> if called within <see cref="InitializeAsync(CancellationToken)"/>.</param>
    /// <inheritdoc cref="CheckState(bool)"/>
    private object? CheckStateAndHandleError(Func<object?> getValue, bool isInitializing)
    {
        try
        {
            CheckState(isInitializing: isInitializing);

            return getValue();
        }
        catch (ErrorsCountExceededException)
        {
            throw;
        }
        catch (InvalidOperationException)
        {
            throw;
        }
        catch (Exception)
        {
            if (!isInitializing)
            {
                ComObjectErrorsCount++;
            }

            throw;
        }
    }

    /// <summary>
    /// Must to be called at the start of any operation except <see cref="DisposeAsync"/>.
    /// </summary>
    private void StopDisposeTimer() =>
        _disposeTokenSource.CancelAfter(Timeout.InfiniteTimeSpan);

    private void StartDisposeTimer(out TimeSpan disposeTimeout)
    {
        disposeTimeout = DefaultDisposeTimeout;
        _disposeTokenSource.CancelAfter(DefaultDisposeTimeout);
    }

    private async ValueTask WaitDisposeTimerAsync()
    {
        try
        {
            await Task.Delay(Timeout.InfiniteTimeSpan, _disposeTokenSource.Token).ConfigureAwait(false);
        }
        catch (TaskCanceledException)
        {
            // Supress exception
        }
    }

    private void ReleaseComObject()
    {
        if (_comObject is not null)
        {
            _ = Marshal.FinalReleaseComObject(_comObject);
        }

        _isInitialized = false;
        _comObject = null;
        _comObjectType = null;

        GC.Collect();
        GC.WaitForPendingFinalizers();
    }
}

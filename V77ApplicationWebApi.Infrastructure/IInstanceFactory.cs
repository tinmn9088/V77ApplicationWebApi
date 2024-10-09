using System;
using V77ApplicationWebApi.Infrastructure.Exceptions;

namespace V77ApplicationWebApi.Infrastructure;

internal interface IInstanceFactory
{
    /// <summary>
    /// Gets the type associated with the specified program identifier (ProgID).
    /// </summary>
    /// <param name="progID">The ProgID of the type to get.</param>
    /// <returns>The type associated with the specified ProgID.</returns>
    /// <exception cref="FailedToGetTypeFromProgIDException">The specified ProgID is not registered.</exception>
    Type GetTypeFromProgID(string progID);

    /// <summary>
    /// Creates an instance of the specified type using that type's default constructor.
    /// </summary>
    /// <param name="type">The type of object to create.</param>
    /// <returns>A reference to the newly created object.</returns>
    /// <exception cref="FailedToCreateInstanceException">If an error occurred.</exception>
    object? CreateInstance(Type type);
}

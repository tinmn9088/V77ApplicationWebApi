using System;
using V77ApplicationWebApi.Infrastructure.Exceptions;

namespace V77ApplicationWebApi.Infrastructure;

public sealed class InstanceFactory : IInstanceFactory
{
    public Type GetTypeFromProgID(string progID)
    {
        try
        {
            return Type.GetTypeFromProgID(progID, throwOnError: true);
        }
        catch (Exception ex)
        {
            throw new FailedToGetTypeFromProgIDException(progID, ex);
        }
    }

    public object CreateInstance(Type type)
    {
        try
        {
            return Activator.CreateInstance(type);
        }
        catch (Exception ex)
        {
            throw new FailedToCreateInstanceException(type, ex);
        }
    }
}

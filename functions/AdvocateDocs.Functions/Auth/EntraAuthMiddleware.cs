using System.Net;
using System.Security.Claims;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Azure.Functions.Worker.Middleware;
using Microsoft.AspNetCore.Http;

namespace AdvocateDocs.Functions.Auth;

/// <summary>
/// Requires a valid Entra token on every HTTP function except those listed as anonymous; stores the caller's
/// principal in <see cref="FunctionContext.Items"/> under <see cref="PrincipalKey"/>.
/// </summary>
public sealed class EntraAuthMiddleware(ITokenValidator validator) : IFunctionsWorkerMiddleware
{
    public const string PrincipalKey = "CallerPrincipal";

    /// <summary>Functions reachable without a token (health probe only).</summary>
    private static readonly HashSet<string> AnonymousFunctions = new(StringComparer.OrdinalIgnoreCase) { "Health" };

    public async Task Invoke(FunctionContext context, FunctionExecutionDelegate next)
    {
        var httpContext = context.GetHttpContext();
        if (httpContext is null || AnonymousFunctions.Contains(context.FunctionDefinition.Name))
        {
            await next(context);
            return;
        }

        var result = await validator.ValidateAsync(httpContext.Request.Headers.Authorization, context.CancellationToken);
        if (!result.IsSuccess)
        {
            httpContext.Response.StatusCode = (int)HttpStatusCode.Unauthorized;
            await httpContext.Response.WriteAsJsonAsync(new { error = result.Error }, context.CancellationToken);
            return;
        }

        context.Items[PrincipalKey] = result.Value!;
        await next(context);
    }

    /// <summary>The validated caller of the current invocation.</summary>
    public static ClaimsPrincipal GetCaller(FunctionContext context) => (ClaimsPrincipal)context.Items[PrincipalKey];
}

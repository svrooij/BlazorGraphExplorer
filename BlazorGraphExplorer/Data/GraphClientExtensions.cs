using System;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading;
using System.Threading.Tasks;
using Azure.Core;
using MsalAuth = Microsoft.AspNetCore.Components.WebAssembly.Authentication;
using Microsoft.Authentication.WebAssembly.Msal.Models;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Graph;
using Microsoft.Kiota.Abstractions;
using KiotaAuth = Microsoft.Kiota.Abstractions.Authentication;

/// <summary>
/// Adds services and implements methods to use Microsoft Graph SDK.
/// </summary>
internal static class GraphClientExtensions
{
    /// <summary>
    /// Extension method for adding the Microsoft Graph SDK to IServiceCollection.
    /// </summary>
    /// <param name="services"></param>
    /// <param name="scopes">The MS Graph scopes to request</param>
    /// <returns></returns>
    public static IServiceCollection AddMicrosoftGraphClient(this IServiceCollection services, params string[] scopes)
    {
        services.Configure<MsalAuth.RemoteAuthenticationOptions<MsalProviderOptions>>(options =>
        {
            foreach (var scope in scopes)
            {
                options.ProviderOptions.AdditionalScopesToConsent.Add(scope);
            }
        });

        services.AddScoped<KiotaAuth.IAuthenticationProvider, GraphAuthenticationProvider>();
        //services.AddHttpClient();
        services.AddHttpClient(nameof(GraphClientExtensions));


        services.AddScoped(sp => new GraphServiceClient(
                  sp.GetRequiredService<IHttpClientFactory>().CreateClient(nameof(GraphClientExtensions)),
                  sp.GetRequiredService<KiotaAuth.IAuthenticationProvider>())
        );
        return services;
    }

    /// <summary>
    /// Implements IAuthenticationProvider interface.
    /// Tries to get an access token for Microsoft Graph.
    /// </summary>
    private class GraphAuthenticationProvider : KiotaAuth.IAuthenticationProvider
    {
        public GraphAuthenticationProvider(MsalAuth.IAccessTokenProvider provider)
        {
            Provider = provider;
        }

        public MsalAuth.IAccessTokenProvider Provider { get; }

        public async Task AuthenticateRequestAsync(RequestInformation request, Dictionary<string, object>? additionalAuthenticationContext = null, CancellationToken cancellationToken = default)
        {
            var result = await Provider.RequestAccessToken(new MsalAuth.AccessTokenRequestOptions()
            {
                // TODO: Get correct scopes based on request url
                Scopes = new[] { "https://graph.microsoft.com/User.Read" },
            });
            if (result.TryGetToken(out var token))
            {
                request.Headers.Add("Authorization", $"Bearer {token.Value}");
            }
        }
    }
}

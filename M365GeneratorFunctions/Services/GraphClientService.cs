// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

using Azure.Identity;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using Microsoft.Azure.Services.AppAuthentication;
using System.Net.Http.Headers;

namespace M365GeneratorFunctions.Services
{
    public class GraphClientService : IGraphClientService
    {
        private readonly IConfiguration _config;
        private readonly ILogger _logger;
        private GraphServiceClient? _appGraphClient;

        public GraphClientService(IConfiguration config, ILoggerFactory loggerFactory)
        {
            _config = config;
            _logger = loggerFactory.CreateLogger<GraphClientService>();
        }

        public GraphServiceClient? GetUserGraphClient(string userAssertion)
        {
            var tenantId = _config["tenantId"];
            var clientId = _config["apiClientId"];
            var clientSecret = _config["apiClientSecret"];

            if (string.IsNullOrEmpty(tenantId) ||
                string.IsNullOrEmpty(clientId) ||
                string.IsNullOrEmpty(clientSecret))
            {
                _logger.LogError("Required settings missing: 'tenantId', 'apiClientId', and 'apiClientSecret'.");
                return null;
            }

            var onBehalfOfCredential = new OnBehalfOfCredential(
                tenantId, clientId, clientSecret, userAssertion);

            return new GraphServiceClient(onBehalfOfCredential);
        }

        public async Task<GraphServiceClient> GetUserGraphClient()
{
    var azureServiceTokenProvider = new AzureServiceTokenProvider();
    string accessToken = await azureServiceTokenProvider
        .GetAccessTokenAsync("https://graph.microsoft.com/");

    var graphServiceClient = new GraphServiceClient(
        new DelegateAuthenticationProvider((requestMessage) =>
    {
        requestMessage
            .Headers
            .Authorization = new AuthenticationHeaderValue("bearer", accessToken);

        return Task.CompletedTask;
    }));

    return graphServiceClient;
}

        public GraphServiceClient? GetAppGraphClient()
        {
            if (_appGraphClient == null)
            {
                var tenantId = _config["tenantId"];
                var clientId = _config["webhookClientId"];
                var clientSecret = _config["webhookClientSecret"];

                if (string.IsNullOrEmpty(tenantId) ||
                    string.IsNullOrEmpty(clientId) ||
                    string.IsNullOrEmpty(clientSecret))
                {
                    _logger.LogError("Required settings missing: 'tenantId', 'webhookClientId', and 'webhookClientSecret'.");
                    return null;
                }

                var clientSecretCredential = new ClientSecretCredential(
                    tenantId, clientId, clientSecret);

                _appGraphClient = new GraphServiceClient(clientSecretCredential);
            }

            return _appGraphClient;
        }
    }
}

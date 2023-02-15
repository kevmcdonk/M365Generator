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

        public GraphServiceClient GetUserGraphClient()
        {
            if (_appGraphClient == null)
            {
            var tenantId = _config["AZURE_TENANTID"];
            var clientId = _config["AZURE_CLIENT_ID"];
            var username = _config["AZURE_USERNAME"];
            var password = _config["AZURE_PASSWORD"];

            var creds = new UsernamePasswordCredential(username, password, tenantId, clientId);
            _appGraphClient = new GraphServiceClient(creds);
            }
            return _appGraphClient;
        }

        public GraphServiceClient? GetAppGraphClient()
        {
            if (_appGraphClient == null)
            {
                var tenantId = _config["AZURE_TENANTID"];
                var clientId = _config["AZURE_CLIENT_ID"];
                var clientSecret = _config["apiClientSecret"];

                var creds = new ClientSecretCredential(tenantId, clientId, clientSecret);
                _appGraphClient = new GraphServiceClient(creds);
            }

            return _appGraphClient;
        }

        public GraphServiceClient? GetUserGraphClient(string[] scopes)
        {
            if (_appGraphClient == null)
            {
                var tenantId = _config["tenantId"];
                var clientId = _config["apiClientId"];
                var clientSecret = _config["apiClientSecret"];

                if (string.IsNullOrEmpty(tenantId) ||
                    string.IsNullOrEmpty(clientId) ||
                    string.IsNullOrEmpty(clientSecret))
                {
                    _logger.LogError("Required settings missing: 'tenantId', 'webhookClientId', and 'webhookClientSecret'.");
                    return null;
                }

                Func<DeviceCodeInfo, CancellationToken, Task> callback = (code, cancellation) => {
                    Console.WriteLine(code.Message);
                    return Task.FromResult(0);
                };

                var clientSecretCredential = new ClientSecretCredential(
                    tenantId, clientId, clientSecret);

                    var deviceCodeCredential = new DeviceCodeCredential(callback, tenantId, clientId, null);

                _appGraphClient = new GraphServiceClient(deviceCodeCredential, scopes);
            }

            return _appGraphClient;
        }
    }
}

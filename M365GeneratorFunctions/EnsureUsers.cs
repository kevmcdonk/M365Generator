// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

using System.Net;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Azure.Functions.Worker.Http;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using M365GeneratorFunctions.Services;

namespace M365GeneratorFunctions
{
    public class EnsureUsers
    {
        // private readonly ITokenValidationService _tokenValidationService;
        private readonly IGraphClientService _graphClientService;
        private readonly ILogger _logger;

        public EnsureUsers(
            // ITokenValidationService tokenValidationService,
            IGraphClientService graphClientService,
            ILoggerFactory loggerFactory)
        {
            // _tokenValidationService = tokenValidationService;
            _graphClientService = graphClientService;
            _logger = loggerFactory.CreateLogger<EnsureUsers>();
        }

        [Function("EnsureUsers")]
        public async Task<HttpResponseData> RunAsync(
            [HttpTrigger(AuthorizationLevel.Function, "get")] HttpRequestData req)
        {
            _logger.LogInformation("EnsureUsers function triggered.");

            //var graphClient = await _graphClientService.GetUserGraphClient();
            var scopes = new[] { "User.Read" };
            var graphClient = _graphClientService.GetAppGraphClient();
            
            if (graphClient == null)
            {
                _logger.LogError("Could not create a Graph client for the user");
                return req.CreateResponse(HttpStatusCode.InternalServerError);
            }
            
            
            int userFoundCount = 0;
            bool blueyFound = await EnsureUser("bluey@mcd79.com", graphClient);
            if (blueyFound) {
                userFoundCount++;
            } else {
                // Create the user
            }

            bool kwaziiFound = await EnsureUser("kwazii@mcd79.com", graphClient);
            if (kwaziiFound) {
                userFoundCount++;
            } else {
                // Create the user
            }

            bool scoopFound = await EnsureUser("scoop@mcd79.com", graphClient);
            if (scoopFound) {
                userFoundCount++;
            } else {
                // Create the user
            }

            bool stickFound = await EnsureUser("stick@mcd79.com", graphClient);
            if (stickFound) {
                userFoundCount++;
            } else {
                // Create the user
            }

            if (userFoundCount == 4)
            {
                var response = req.CreateResponse(HttpStatusCode.OK);
                // Return the message in the response
                response.WriteString("All users found");
                return response;
            }
            else {
                var response = req.CreateResponse(HttpStatusCode.InternalServerError);
                // Return the message in the response
                response.WriteString("There were " + userFoundCount.ToString() + " users found");
                return response;
            }

            return req.CreateResponse(HttpStatusCode.NoContent);
        }

        private async Task<bool> EnsureUser(string emailAddress, GraphServiceClient graphClient) {
            var queryOptions = new List<QueryOption>()
            {
                new QueryOption("ConsistencyLevel", "eventual"),
                new QueryOption("$count", "true")
            };

            var users = await graphClient.Users
	        .Request( queryOptions )
            .Filter("mail eq '" + emailAddress + "'")
            .GetAsync();

            if (users.CurrentPage.Count > 0)
            {
                return true;
            }

            return false;
        }
    }
}

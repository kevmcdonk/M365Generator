// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

using System.Net;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Azure.Functions.Worker.Http;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using System.Net.Http;
using System.Net.Http.Headers;
using M365GeneratorFunctions.Services;

namespace M365GeneratorFunctions
{
    public class EnsureTeams
    {
        // private readonly ITokenValidationService _tokenValidationService;
        private readonly IGraphClientService _graphClientService;
        private readonly ILogger _logger;
        private readonly IConfiguration _config;

        public EnsureTeams(
            IConfiguration config,
            IGraphClientService graphClientService,
            ILoggerFactory loggerFactory)
        {
            _config = config;
            // _tokenValidationService = tokenValidationService;
            _graphClientService = graphClientService;
            _logger = loggerFactory.CreateLogger<EnsureUsers>();
        }

        [Function("EnsureTeams")]
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
            
            int teamFoundCount = 0;
            
            var queryOptions = new List<QueryOption>()
            {
                new QueryOption("ConsistencyLevel", "eventual"),
                new QueryOption("$count", "true")
            };

            var teams = await graphClient.Groups
	        .Request( queryOptions )
            .Filter("resourceProvisioningOptions/Any(x:x eq 'Team')")
            .GetAsync();
            
            teamFoundCount = teams.CurrentPage.Count;

            if (teamFoundCount < 30)
            {
                var response = req.CreateResponse(HttpStatusCode.InternalServerError);
                // Return the message in the response
                response.WriteString("There were only " + teamFoundCount.ToString() + " teams found");
                return response;

                //TODO: Create the Teams
            }
            else {
                var response = req.CreateResponse(HttpStatusCode.OK);
                // Return the message in the response
                response.WriteString("Enough Teams found");
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

        private async Task<string[]> GetRandomTeamNames(CancellationToken cancellationToken = default) {
            HttpClient Client = new HttpClient();
            var openAIKey = _config["AZURE_PASSWORD"];
            Client.DefaultRequestHeaders.Add("User-Agent", "OpenAI-DotNet");
            Client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", openAIKey);

            var bodyText = "{\"prompt\": \"Generate a list of teams for Microsoft Teams projects\",  \"max_tokens\": 60,\"temperature\": 0.8,\"frequency_penalty\": 0,\"presence_penalty\": 0,\"top_p\": 1,\"best_of\": 1,\"stop\": null}";
            //var jsonContent = JsonSerializer.Serialize(completionRequest, Api.JsonSerializationOptions);
            var response = await Client.PostAsync("https://m365generator.openai.azure.com/openai/deployments/text-davinci-002/completions?api-version=2022-12-01", bodyText, cancellationToken).ConfigureAwait(false);
            var responseAsString = await response.ReadAsStringAsync(cancellationToken).ConfigureAwait(false);
            //return DeserializeResult(response, responseAsString);
            return new string[] { "test", "test2" };
            /*
            curl https://m365generator.openai.azure.com/openai/deployments/text-davinci-002/completions?api-version=2022-12-01 \
  -H "Content-Type: application/json" \
  -H "api-key: YOUR_API_KEY" \
  -d '{
  "prompt": "Generate a list of teams for Microsoft Teams projects\n\n-IT support team\n-Marketing team\n-Sales team\n-Human Resources team\n-Finance team\n-Legal team\n-Information Technology team\n-Product Development team\n-Customer Service team",
  "max_tokens": 60,
  "temperature": 0.8,
  "frequency_penalty": 0,
  "presence_penalty": 0,
  "top_p": 1,
  "best_of": 1,
  "stop": null
}'*/
        }
    }
}

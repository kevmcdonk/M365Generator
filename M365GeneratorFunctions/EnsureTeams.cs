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
using System.Text;
using Newtonsoft.Json;

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
                string[] teamNames = await GetRandomTeamNames();
                foreach (string teamName in teamNames) {
                    if (teamName != "") {
                        string teamNameToCreate = teamName;
                        if (teamName.StartsWith("-")) {
                            teamNameToCreate = teamName.Substring(1,teamName.Length-1);
                        }
                        await CreateTeam(teamNameToCreate, graphClient);
                    }
                }
               
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

        private async Task CreateTeam(string teamName, GraphServiceClient graphClient) {

            var additionalData = new Dictionary<string, object>()
    {
        {"owners@odata.bind", new List<string>()},
        {"members@odata.bind", new List<string>()}
    };
(additionalData["members@odata.bind"] as List<string>).Add("https://graph.microsoft.com/v1.0/users/48e8a1ab-0d3a-4f9b-b200-e9e9d1437a2b");
(additionalData["owners@odata.bind"] as List<string>).Add("https://graph.microsoft.com/v1.0/users/48e8a1ab-0d3a-4f9b-b200-e9e9d1437a2b");

            var newGroup = new Group 
            {
                DisplayName = teamName,
                MailNickname = teamName.Replace(" ",""),
                Description = teamName,
                Visibility = "Private",
                GroupTypes = new List<String>(){ "Unified"},
                MailEnabled=true,
                SecurityEnabled=false,
                AdditionalData = additionalData
            };

            Group createdGroup = await graphClient.Groups
            .Request()
            .AddAsync(newGroup);

            var team = new Team
{
	MemberSettings = new TeamMemberSettings
	{
		AllowCreatePrivateChannels = true,
		AllowCreateUpdateChannels = true
	},
	MessagingSettings = new TeamMessagingSettings
	{
		AllowUserEditMessages = true,
		AllowUserDeleteMessages = true
	},
	FunSettings = new TeamFunSettings
	{
		AllowGiphy = true,
		GiphyContentRating = GiphyRatingType.Strict
	}
};
Thread.Sleep(15000);
await graphClient.Groups[createdGroup.Id].Team
	.Request()
	.PutAsync(team);
        }

        private async Task<string[]> GetRandomTeamNames(CancellationToken cancellationToken = default) {
            HttpClient Client = new HttpClient();
            var openAIKey = _config["AZURE_OPENAI_ID"];
            Client.DefaultRequestHeaders.Add("User-Agent", "OpenAI-DotNet");
            Client.DefaultRequestHeaders.Add("api-key", openAIKey);
            Client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));//ACCEPT header


            string bodyText = "{\"prompt\": \"Generate a list of teams for Microsoft Teams projects\",  \"max_tokens\": 60,\"temperature\": 0.8,\"frequency_penalty\": 0,\"presence_penalty\": 0,\"top_p\": 1,\"best_of\": 1,\"stop\": null}";
            //var jsonContent = JsonSerializer.Serialize(completionRequest, Api.JsonSerializationOptions);
            HttpContent content = new StringContent(bodyText);
            string endPoint = "https://m365generator.openai.azure.com/openai/deployments/text-davinci-002/completions?api-version=2022-12-01"; //_config["AZURE_OPENAI_ENDPOINT"];
            //var response = await Client.PostAsync(endPoint, content, cancellationToken).ConfigureAwait(false);
            //var responseText = await response.Content.ReadAsStringAsync();


            
            HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Post, endPoint);
            request.Content = new StringContent(bodyText, Encoding.UTF8, "application/json");
            var response = await Client.SendAsync(request);
            var responseAsString = await response.Content.ReadAsStringAsync();
            dynamic stuff = JsonConvert.DeserializeObject(responseAsString);
            return stuff.choices[0].text.ToString().Split("\n");
            //return DeserializeResult(response, responseAsString);
            //return new string[] { "test", "test2" };
            /*
            "{\"id\":\"cmpl-6l30pLDiVwr6mhvjvS3F0Rk90PEwS\",\"object\":\"text_completion\",\"created\":1676671015,\"model\":\"text-davinci-002\",\"choices\":[{\"text\":\"\\n\\n1. Engineering\\n2. Marketing\\n3. Sales\\n4. Customer Support\\n5. Human Resources\\n6. Accounting\\n7. IT\\n8. Facilities\\n9. Security\\n10. Legal\",\"index\":0,\"logprobs\":null,\"finish_reason\":\"stop\"}],\"usage\":{\"prompt_tokens\":10,\"completion_tokens\":43,\"total_tokens\":53}}\n"

}'*/
        }
    }
}

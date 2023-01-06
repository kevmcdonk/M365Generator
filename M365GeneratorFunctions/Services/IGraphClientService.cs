// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

using Microsoft.Graph;

namespace M365GeneratorFunctions.Services
{
    public interface IGraphClientService
    {
        public GraphServiceClient? GetUserGraphClient(string userAssertion);
        public Task<GraphServiceClient> GetUserGraphClient();
        public GraphServiceClient? GetAppGraphClient();
    }
}

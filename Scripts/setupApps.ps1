# Requires CLI for Microsoft 365 to be installed - https://pnp.github.io/cli-microsoft365
$ApplicationName = 'M365 Generator Azure Function'

# m365 aad app add --name "$ApplicationName Test App" --redirectUris 'http://localhost:8080' --platform spa --apisDelegated 'https://graph.microsoft.com/Mail.Read' --save
# m365 aad app add --name "$ApplicationName" --withSecret --apisDelegated 'https://graph.microsoft.com/User.Read.All,https://graph.microsoft.com/Mail.Read' --save

#Set permissions for managed identity - https://powers-hell.com/2022/09/12/authenticate-to-graph-in-azure-functions-with-managed-identities/
Connect-AzAccount -Tenant b618d3d2-c131-42fd-a629-3711d84af275 -Subscription c6f9fb91-7294-4aad-8826-219820a8d8bb
$token = (Get-AzAccessToken -ResourceUrl "https://graph.microsoft.com").Token

$baseUri = 'https://graph.microsoft.com/v1.0/servicePrincipals'
$graphAppId = '00000003-0000-0000-c000-000000000000'
$spSearchFiler = '"displayName:{0}" OR "appId:{1}"' -f $ApplicationName, $graphAppId
$msiParams = @{
    Method  = 'Get'
    Uri     = '{0}?$search={1}' -f $baseUri, $spSearchFiler
    Headers = @{Authorization = "Bearer $Token"; ConsistencyLevel = "eventual" }
}
$spList = (Invoke-RestMethod @msiParams).Value
$msiId = ($spList | Where-Object { $_.displayName -eq $applicationName }).Id
$graphId = ($spList | Where-Object { $_.appId -eq $graphAppId }).Id

$roles = @(
    "User.Read.All",
    "Mail.ReadWrite"
)

$graphRoleParams = @{
    Method  = 'Get'
    Uri     = "$baseUri/$($GraphId)/appRoles"
    Headers = @{Authorization = "Bearer $Token"; ConsistencyLevel = "eventual" }
}
$graphRoles = (Invoke-RestMethod @graphRoleParams).Value | 
        Where-Object {$_.value -in $roles -and $_.allowedMemberTypes -Contains "Application"} |
        Select-Object allowedMemberTypes, id, value

$baseUri = 'https://graph.microsoft.com/v1.0/servicePrincipals'

foreach ($role in $graphRoles) {
    $postBody = @{
        "principalId" = $msiId
        "resourceId"  = $graphId
        "appRoleId"   = $role.Id
    }
    $restParams = @{
        Method      = "Post"
        Uri         = "$baseUri/$($graphId)/appRoleAssignedTo"
        Body        = $postBody | ConvertTo-Json
        Headers     = @{Authorization = "Bearer $token" }
        ContentType = 'Application/Json'
    }
    $roleRequest = Invoke-RestMethod @restParams
    $roleRequest
}









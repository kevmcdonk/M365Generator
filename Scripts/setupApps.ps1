# Requires CLI for Microsoft 365 to be installed - https://pnp.github.io/cli-microsoft365
m365 aad app add --name 'M365 Generator Azure Function Test App' --redirectUris 'http://localhost:8080' --platform spa --apisDelegated 'https://graph.microsoft.com/Mail.Read' --save
m365 aad app add --name 'M365 Generator Azure Function' --withSecret --apisDelegated 'https://graph.microsoft.com/User.Read.All,https://graph.microsoft.com/Mail.Read' --save

# Grant permissions
m365 aad approleassignment add --appDisplayName 'M365 Generator Azure Function Test App' --resource "Microsoft Graph" --scope "Mail.Read,Mail.Send"



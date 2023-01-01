---
ms.localizationpriority: medium
---

<!-- markdownlint-disable MD002 MD041 -->

This tutorial teaches you how to build an Azure Function that uses the Microsoft Graph API to retrieve calendar information for a user.

> [!TIP]
> If you prefer to just download the completed tutorial, you can download or clone the [GitHub repository](https://github.com/microsoftgraph/msgraph-training-azurefunction-csharp). See the README file in the **demo** folder for instructions on configuring the app with an app ID and secret.

## Prerequisites

Before you start this tutorial, you should have the following tools installed on your development machine.

- [.NET Core SDK](https://dotnet.microsoft.com/download)
- [Azure Functions Core Tools](https://docs.microsoft.com/azure/azure-functions/functions-run-local)
- [Azure CLI](https://docs.microsoft.com/cli/azure/install-azure-cli)
- [ngrok](https://ngrok.com/)

You should also have a Microsoft work or school account, with access to a global administrator account in the same organization. If you don't have a Microsoft account, you can [sign up for the Microsoft 365 Developer Program](https://developer.microsoft.com/microsoft-365/dev-program) to get a free Office 365 subscription.

> [!NOTE]
> This tutorial was written with the following versions of the above tools. The steps in this guide may work with other versions, but that has not been tested.
>
> - .NET Core SDK 5.0.203
> - Azure Functions Core Tools 3.0.3442
> - Azure CLI 2.23.0
> - ngrok 2.3.40

## Feedback

Please provide any feedback on this tutorial in the [GitHub repository](https://github.com/microsoftgraph/msgraph-training-azurefunction-csharp).

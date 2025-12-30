# Agent365 Notification Agent

A Microsoft Agent 365 sample agent built with .NET 8 that demonstrates how to handle notifications from multiple Microsoft 365 applications. This agent can respond to messages in Microsoft Teams, reply to comments in Word documents, and respond to email messages.

## Overview

This project showcases the Microsoft Agent 365 notification system, which enables agents to receive and respond to events from across the Microsoft 365 ecosystem. The agent uses Azure OpenAI for natural language processing and MCP (Model Context Protocol) tool servers for interacting with Microsoft 365 services.

## Capabilities

| Channel | Capability |
|---------|------------|
| **Microsoft Teams** | Respond to direct messages and mentions in conversations |
| **Microsoft Word** | Respond to comments where the agent is mentioned |
| **Microsoft Outlook** | Reply to emails where the agent is addressed or mentioned |

## Architecture

```
┌─────────────────────────────────────────────────────────────────┐
│                    Microsoft 365 Applications                   │
│           (Teams, Word, Outlook, Excel, PowerPoint)             │
└──────────────────────────┬──────────────────────────────────────┘
                           │ Notifications
                           ▼
┌─────────────────────────────────────────────────────────────────┐
│                     Agent Framework                             │
│  ┌─────────────────┐  ┌─────────────────┐  ┌─────────────────┐  │
│  │ OnAgenticWord   │  │ OnAgenticEmail  │  │ OnActivity      │  │
│  │ Notification    │  │ Notification    │  │ (Teams/General) │  │
│  └────────┬────────┘  └────────┬────────┘  └────────┬────────┘  │
│           │                    │                    │           │
│           └────────────────────┴────────────────────┘           │
│                              │                                  │
│                    ┌─────────▼─────────┐                        │
│                    │   Azure OpenAI    │                        │
│                    │   (Chat Client)   │                        │
│                    └─────────┬─────────┘                        │
│                              │                                  │
│                    ┌─────────▼─────────┐                        │
│                    │   MCP Tool        │                        │
│                    │   Servers         │                        │
│                    └───────────────────┘                        │
└─────────────────────────────────────────────────────────────────┘
```

## Notification Handlers

The agent registers three main handlers in [MyAgent.cs](notification-agent/Agent/MyAgent.cs):

### OnAgenticWordNotification

```csharp
this.OnAgenticWordNotification(HandleWordCommentNotificationAsync, autoSignInHandlers: new[] { AgenticIdAuthHandler });
```

This handler is invoked when the agent is mentioned in a comment within a Word document. When triggered:

1. **Receives notification data** - The `AgentNotificationActivity` contains a `WpxCommentNotification` with details about the comment, including the `CommentId` and document information.

2. **Extracts document content** - The handler retrieves the document's content URL from the activity attachments and uses the `WordGetDocumentContent` MCP tool to fetch the full document text.

3. **Processes the comment** - Using Azure OpenAI, the agent analyzes the document context and the specific text the comment refers to, then formulates an appropriate response.

4. **Replies via MCP** - Instead of sending a direct activity response, the agent uses the Word MCP Server tools to post a reply directly to the comment thread in the document.

**Example flow:**
- User mentions `@Agent` in a Word comment asking "Can you summarize this paragraph?"
- Agent receives the notification with document URL and comment ID
- Agent fetches document content and identifies the referenced text
- Agent generates a summary and posts it as a reply to the comment

### OnAgenticEmailNotification

```csharp
this.OnAgenticEmailNotification(HandleEmailNotificationAsync, autoSignInHandlers: new[] { AgenticIdAuthHandler });
```

This handler is invoked when the agent receives an email where it is mentioned or addressed. When triggered:

1. **Receives email notification** - The `AgentNotificationActivity` contains an `EmailNotification` object with the email's `Id` and `ConversationId`.

2. **Extracts email content** - The handler retrieves the email content from `turnContext.Activity.Text`.

3. **Generates a response** - Using Azure OpenAI, the agent processes the email content and formulates an appropriate reply.

4. **Sends the reply** - The agent uses the `ReplyToMessageAsync` MCP tool to send an HTML-formatted reply to the original email.

**Example flow:**
- User sends an email to the agent asking "What's the status of Project X?"
- Agent receives the notification with the email ID and content
- Agent processes the request and generates a response
- Agent sends the reply email using the email MCP tools

### OnActivity (Teams/General Messages)

```csharp
this.OnActivity(ActivityTypes.Message, OnMessageAsync, autoSignInHandlers: new[] { AgenticIdAuthHandler });
```

This is the general message handler for Teams conversations and other channels. It:

1. Maintains conversation threads for context continuity
2. Processes user messages using Azure OpenAI
3. Responds directly in the conversation channel

> **Note:** This handler is registered last to ensure it doesn't intercept notifications meant for the Word and Email handlers.

## Prerequisites

- .NET 8.0 SDK
- Azure OpenAI service with a deployed model
- Microsoft Agent 365 Frontier preview program access
- Configured MCP tool servers for Word and Email integration

## Configuration

The agent requires the following configuration in `appsettings.json` or user secrets:

```json
{
  "AIServices": {
    "AzureOpenAI": {
      "Endpoint": "https://your-openai-instance.openai.azure.com/",
      "ApiKey": "your-api-key",
      "DeploymentName": "your-deployment-name"
    }
  }
}
```

## Project Structure

```
notification-agent/
├── Agent/
│   └── MyAgent.cs              # Main agent implementation with notification handlers
├── Tools/
│   └── DateTimeFunctionTool.cs # Custom tool for date/time operations
├── Program.cs                  # Application entry point and DI configuration
└── AgentFrameworkNotificationAgent.csproj
```

## Running the Agent

1. Configure your Azure OpenAI credentials
2. Set up MCP tool server configuration for Word and Email services
3. Run the application:

```bash
cd notification-agent
dotnet run
```

The agent will start listening on `http://localhost:3978` in development mode.

## Key Dependencies

- `Microsoft.Agents.A365.Notifications` - Agent 365 notification handling
- `Microsoft.Agents.A365.Tooling.Extensions.AgentFramework` - MCP tool integration
- `Microsoft.Agents.Builder` - Agent Framework core
- `Azure.AI.OpenAI` - Azure OpenAI client
- `Microsoft.Extensions.AI` - AI abstractions and extensions

## Documentation

For more information about Microsoft Agent 365 notifications, see the [official documentation](https://learn.microsoft.com/en-us/microsoft-agent-365/developer/notification?tabs=dotnet).

## License

Copyright (c) Microsoft Corporation. All rights reserved.
Licensed under the MIT License.

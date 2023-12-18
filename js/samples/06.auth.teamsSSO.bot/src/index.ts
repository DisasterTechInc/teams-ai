/* eslint-disable @typescript-eslint/no-unused-vars */
// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

// Import required packages
import { config } from 'dotenv';
import * as path from 'path';
import * as restify from 'restify';

import {
    Application,
    preview,
    AI
} from '@microsoft/teams-ai';

// Import required bot services.
// See https://aka.ms/bot-services to learn more about the different parts of a bot.
import {
    ActivityTypes,
    CloudAdapter,
    ConfigurationBotFrameworkAuthentication,
    ConfigurationServiceClientCredentialFactory,
    MemoryStorage,
    TurnContext
} from 'botbuilder';

const { AssistantsPlanner } = preview;

// Read botFilePath and botFileSecret from .env file.
const ENV_FILE = path.join(__dirname, '..', '.env');
config({ path: ENV_FILE });

const botFrameworkAuthentication = new ConfigurationBotFrameworkAuthentication(
    {},
    new ConfigurationServiceClientCredentialFactory({
        MicrosoftAppId: process.env.BOT_ID,
        MicrosoftAppPassword: process.env.BOT_PASSWORD,
        MicrosoftAppType: 'MultiTenant'
    })
);

// Create adapter.
// See https://aka.ms/about-bot-adapter to learn more about how bots work.
const adapter = new CloudAdapter(botFrameworkAuthentication);

// Catch-all for errors.
const onTurnErrorHandler = async (context: any, error: any) => {
    // This check writes out errors to console log .vs. app insights.
    // NOTE: In production environment, you should consider logging this to Azure
    //       application insights.
    console.error(`\n [onTurnError] unhandled error: ${error.toString()}`);

    // Send a trace activity, which will be displayed in Bot Framework Emulator
    await context.sendTraceActivity(
        'OnTurnError Trace',
        `${error.toString()}`,
        'https://www.botframework.com/schemas/error',
        'TurnError'
    );

    // Send a message to the user
    await context.sendActivity('The bot encountered an error or bug.');
    await context.sendActivity('To continue to run this bot, please fix the bot source code.');
};

// Set the onTurnError for the singleton CloudAdapter.
adapter.onTurnError = onTurnErrorHandler;

// Create HTTP server.
const server = restify.createServer();
server.use(restify.plugins.bodyParser());

server.listen(process.env.port || process.env.PORT || 3978, () => {
    console.log(`\n${server.name} listening to ${server.url}`);
    console.log('\nGet Bot Framework Emulator: https://aka.ms/botframework-emulator');
    console.log('\nTo test your bot in Teams, sideload the app manifest.json within Teams Apps.');
});

import { ApplicationBuilder, AuthError, TurnState } from '@microsoft/teams-ai';

interface ConversationState {
    count: number;
}
type ApplicationTurnState = TurnState<ConversationState>;

// Create Assistant Planner
const planner = new AssistantsPlanner({
    apiKey: process.env.OPENAI_API_KEY!,
    assistant_id: process.env.ASSISTANT_ID!
});

const aiOptions = {
    planner
}

// Define storage and application
const storage = new MemoryStorage();
const app = new ApplicationBuilder<ApplicationTurnState>()
    .withStorage(storage)
    .withAIOptions(aiOptions)
    .withAuthentication(adapter, {
        settings: {
            graph: {
                scopes: ['User.Read'],
                msalConfig: {
                    auth: {
                        clientId: process.env.AAD_APP_CLIENT_ID!,
                        clientSecret: process.env.AAD_APP_CLIENT_SECRET!,
                        authority: `${process.env.AAD_APP_OAUTH_AUTHORITY_HOST}/${process.env.AAD_APP_TENANT_ID}`,
                    }
                },
                signInLink: `https://${process.env.BOT_DOMAIN}/auth-start.html`,
                endOnInvalidMessage: true
            }
        }
    })
    .build();

// Listen for user to say '/reset' and then delete conversation state
app.message('/reset', async (context: TurnContext, state: ApplicationTurnState) => {
    state.deleteConversationState();
    await context.sendActivity(`Ok I've deleted the current conversation state.`);
});

app.message('/signout', async (context: TurnContext, state: ApplicationTurnState) => {
    await app.authentication.signOutUser(context, state);

    // Echo back users request
    await context.sendActivity(`You have signed out`);
});


app.ai.action('healthcheck', async (context: TurnContext, state: ApplicationTurnState) => {
    
  console.log("health check!!!")

    console.log(state.temp.authTokens['graph']);

    const secret = process.env.DISASTER_GPT_SECRET || ""

    fetch('http://localhost:3001/authcheck', {
        method: "GET",
        headers: {
          'x-botapp': 'true',
          'DISASTER_GPT_SECRET': secret,
          'token': state.temp.authTokens['graph']
        }
    }).then(response => {
      if (!response.ok) {
        throw new Error(`HTTP error! Status: ${response.status}`);
      }
      return response.text();
    })
    .then(async (data: string) => {
      console.log(data); // Handle the response data

      console.log("healthcheck result!",data)

      await context.sendActivity(`Ran a health check with result: ${data}`);
      return 'success'
    })
    .catch(async (error: Error) => {
      await context.sendActivity(`Ran a health check with ERROR: ${JSON.stringify(error)}`);
      console.error('There was an error!', error);
      return 'error'
    });
  return 'ran health check'
});

/*
// Listen for ANY message to be received. MUST BE AFTER ANY OTHER MESSAGE HANDLERS
app.activity(ActivityTypes.Message, async (context: TurnContext, state: ApplicationTurnState) => {
    // Increment count state
    let count = state.conversation.count ?? 0;
    state.conversation.count = ++count;

   
    // Echo back users request
    await context.sendActivity(`[${count}] you said: ${context.activity.text}`);
});
*/

app.authentication.get('graph').onUserSignInSuccess(async (context: TurnContext, state: ApplicationTurnState) => {
    // Successfully logged in
    await context.sendActivity('Successfully logged in.  Please repeat your last message to proceed');
    //await context.sendActivity(`Token string length: ${state.temp.authTokens['graph']!.length}`);
    //await context.sendActivity(`This is what you said before the AuthFlow started: ${context.activity.text}`);
});

app.authentication
    .get('graph')
    .onUserSignInFailure(async (context: TurnContext, _state: ApplicationTurnState, error: AuthError) => {
        // Failed to login
        await context.sendActivity('Failed to login');
        await context.sendActivity(`Error message: ${error.message}`);
    });

// Listen for incoming server requests.
server.post('/api/messages', async (req, res) => {
    // Route received a request to adapter for processing
    await adapter.process(req, res as any, async (context) => {
        // Dispatch to application for routing
        await app.run(context);
    });
});

server.get(
    "/auth-:name(start|end).html",
    restify.plugins.serveStatic({
      directory: path.join(__dirname, "public"),
    })
  );

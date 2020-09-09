// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

// index.js is used to setup and configure your bot

// Import required pckages
const path = require('path');
const express = require('express');

// Import required bot services.
// See https://aka.ms/bot-services to learn more about the different parts of a bot.
const { BotFrameworkAdapter, MemoryStorage, UserState } = require('botbuilder');

// Import bot definitions
const { BotActivityHandler } = require('./botActivityHandler');
const { memory } = require('console');

// Read botFilePath and botFileSecret from .env file.
const ENV_FILE = path.join(__dirname, '.env');
require('dotenv').config({ path: ENV_FILE });

// Create adapter.
// See https://aka.ms/about-bot-adapter to learn more about adapters.
const adapter = new BotFrameworkAdapter({
    appId: process.env.BotId,
    appPassword: process.env.BotPassword
});

adapter.onTurnError = async (context, error) => {
    // This check writes out errors to console log .vs. app insights.
    // NOTE: In production environment, you should consider logging this to Azure
    //       application insights.
    console.error(`  [onTurnError] unhandled error: ${ error }`);

    // Send a trace activity, which will be displayed in Bot Framework Emulator
    await context.sendTraceActivity(
        'OnTurnError Trace',
        `${ error }`,
        'https://www.botframework.com/schemas/error',
        'TurnError'
    );

    // Send a message to the user
    await context.sendActivity('The bot encountered an error or bug.');
    await context.sendActivity('To continue to run this bot, please fix the bot source code.');

    await conversationState.delete(context);
};

// DMM: Added storage and states to help with authentication
const memoryStorage = new MemoryStorage();
const userState = new UserState(memoryStorage);

// Create bot handlers
const botActivityHandler = new BotActivityHandler(userState);

// Create HTTP server.
const server = express();
server.use(express.static('pages'));

const port = process.env.port || process.env.PORT || 3978;
server.listen(port, () => 
    console.log(`\Bot/ME service listening at http://localhost:${port}`)
);

// Listen for incoming requests.
server.post('/api/messages', (req, res) => {
    adapter.processActivity(req, res, async (context) => {
        // Process bot activity
        // console.log(`process activity ${JSON.stringify(context.activity)}`);

        await botActivityHandler.run(context);
    });
});
// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

const {
    TeamsActivityHandler,
    CardFactory,
    ActionTypes
} = require('botbuilder');
const { query } = require('express');
const { GraphClient } = require('./graphClient');

const USER_CONFIGURATION = 'userConfigurationProperty';
const DEBUG = true;
const FULL_DEBUG = false;

//DMM Graph API Client will be added here

class BotActivityHandler extends TeamsActivityHandler {
    constructor(userState) {
        super();
        if(!userState) throw new Error("[BotActivityHandler]: Missing Parameter userState.");

        console.log(`[botActivityHandler::constructor] ConnectionName: ${process.env.connectionName}`);
        this.connectionName = process.env.connectionName;
        this.userState = userState;
        this.userConfigurationState = this.userState.createProperty(USER_CONFIGURATION);

        this.onMessage(async (context, next) => {
            console.log(`running dialog with message activity ${context.activity.text}`);
            if(context.activity.text === "logout") {
                await this.logout(context);
            }
            else if(context.activity.text === "help") {
                await context.sendActivity("Supported Commands: logout, help, tasklists, tasks, test task");
            }
            else if(context.activity.text === "tasklists") {
                const token = await this.getLoginToken(context, this.userState);
                if(!token) {
                    await context.sendActivity("You need to run the Todo Sign In before using this command.");
                }
                else {
                    const graphClient = new GraphClient(token);
                    const taskLists = await graphClient.getTaskList();
                    var taskListNames = "";
                    taskLists.value.forEach(tl => {
                        taskListNames = taskListNames + "<br>" + tl.displayName;
                    });
                    await context.sendActivity(`Your Task Lists are: ${taskListNames}`);
                }
            }
            else if(context.activity.text === "tasks") {
                const token = await this.getLoginToken(context, this.userState);
                if(!token) {
                    await context.sendActivity("You need to run the Todo Sign In before using this command.")
                }
                else {
                    const graphClient = new GraphClient(token);
                    const tasks = await graphClient.getDefaultTasks();
                    var taskNames = "";
                    tasks.value.forEach(t => {
                        taskNames = taskNames + "<br>" + t.title + " - " + t.status;
                    });
                    await context.sendActivity(`Here are you default task list tasks:${ taskNames }`);
                }
                
            }
            else if(context.activity.text == "test task") {
                await context.sendActivity("Can you provide feedback on the document by next week?");
            } 
            await next();
        });
    }

    /**
     * override run
     */
    async run(context) {
        await super.run(context);

        // Save any state changes. The load happened during the execution of the Dialog.
        await this.userState.saveChanges(context);
    }

    async logout(context) {
        const botAdapter = context.adapter;
        await botAdapter.signOutUser(context, process.env.ConnectionName);
        await context.sendActivity('You have been signed out.');
    }

    async handleTeamsMessagingExtensionConfigurationSettings(context, settings) {
        if(settings.state != null) {
            await this.userConfigurationState.set(context, settings.state);
        }
    }

    /* Messaging Extension - Action */
    /* Building a messaging extension action command is a two-step process.
        (1) Define how the messaging extension will look and be invoked in the client.
            This can be done from the Configuration tab, or the Manifest Editor.
            Learn more: https://aka.ms/teams-me-design-action.
        (2) Define how the bot service will respond to incoming action commands.
            Learn more: https://aka.ms/teams-me-respond-action.
        
        NOTE:   Ensure the bot endpoint that services incoming messaging extension queries is
                registered with Bot Framework.
                Learn more: https://aka.ms/teams-register-bot.
    */

    // Invoked when the service receives an incoming action command.
    async handleTeamsMessagingExtensionSubmitAction(context, action) {
        if(DEBUG){
            console.log(`running handleTeamsMessagingExtensionSubmitAction: ${action.commandId}`);
            if(FULL_DEBUG) {
                console.log(`full action: ${JSON.stringify(action)}`);
            }
        }
        
        /* Commands are defined in the manifest file. This can be done using the Configuration tab, editing
           the manifest Json directly, or using the manifest editor in App Studio. This project includes two
           commands to help get you started: createCard and shareMessage.
        */
        switch (action.commandId) {
        case 'newTask':
            return await this.newTaskCommand(context, action);
        case 'login':
            return await this.handleLoginCommand(context, action);
        default:
            throw new Error('NotImplemented');
        }
    }
    /* Messaging Extension - Action */
    async handleLoginCommand(context, action) {
        
        const token = await this.getLoginToken(context, action.state);
        if(FULL_DEBUG) { console.log(`[botActivityHandler::handleLoginCommand]: token ${ token }`); }
        if(!token) {
            return await this.promptForLogin(context, action);
        }

        const heroCard = CardFactory.heroCard("Your Token", token);
        heroCard.content.subTitle = "You are signed in";
        const attachment = { contentType: heroCard.contentType, content: heroCard.content, preview: heroCard };

        return {
            composeExtension: {
                type: 'result', 
                attachmentLayout: 'list',
                attachments: [
                    attachment
                ]
            }
        };
    }



    async getLoginToken(context, state) {
        if(DEBUG) { console.log(`[botActivityHandler::getLoginToken] ConnectionName: ${this.connectionName}`); }

        const magicCode = (state && Number.isInteger(Number(state))) ? state: '';
        const tokenResponse = await context.adapter.getUserToken(context, this.connectionName, magicCode);
        if(FULL_DEBUG) { console.log(`[botActivityHandler::getLoginToken] TokenResponse ${JSON.stringify(tokenResponse)}`); }

        if(!tokenResponse || !tokenResponse.token) {                     
            if(DEBUG) { console.log(`[botActivityHandler::getLoginToken] returning null`); }
            return null;
        }

        if(FULL_DEBUG) { console.log(`[botActivityHandler::getLoginToken] returning ${tokenResponse.token}`); }
        return tokenResponse.token;
    }

    async promptForLogin(context, action) {
        const signInLink = await context.adapter.getSignInLink(context, this.connectionName);
        if(DEBUG) { console.log(`[botActivityHandler::promptForLogin]: Sign In at ${signInLink}`); }

        return {
            composeExtension: {
                type: 'auth',
                suggestedActions: {
                    actions: [{
                        type: 'openUrl',
                        value: signInLink,
                        title: 'Bot Service OAuth'
                    }]
                }
            }
        };
    }

    

    async newTaskCommand(context, action) {
        const token = await this.getLoginToken(context, action.state);
        if(!token) {
            return await this.promptForLogin(context, action);
        }

        const graphClient = new GraphClient(token);

        const newTask = await graphClient.createNewTask({
            title: action.data.title
            , linkedResources: [{
                webUrl: action.messagePayload.linkToMessage,
                applicationName: "Microsof Teams",
                displayName: "Microsoft Teams"
            }]
            , body: {
                content: action.messagePayload.body.content,
                contentType: action.messagePayload.body.contentType
            }
            /* Something is wrong with date format.  May require a second call to graph to set the dueDate? */
            //, dueDateTime: {
            //    dateTime: `${action.data.dueDate}T${action.data.dueDateTime}`,
            //    dateTimeZone: "Eastern Standard Time"
            //}
            , status: action.data.taskStatus
            , importance: action.data.importance
        })
        .catch((err) => {
            console.log(`[botActivityHandler::newTaskCommand] graphClient.createNewTask error ${ JSON.stringify(err) }`);
            return err;
        });

        var heroCard = null;
        console.log(`createNewTask returned ${newTask.statusCode}`);
        if(newTask.statusCode && newTask.statusCode !== 200) {
            heroCard = CardFactory.heroCard(newTask.code, newTask.message);
            heroCard.content.subTitle = `Status Code: ${ newTask.statusCode }`;
        }
        else {
            heroCard = CardFactory.heroCard(action.data.title, action.messagePayload.body.content);
            heroCard.content.subTitle = "Task Created Successfully";
        }
        const attachment = { contentType: heroCard.contentType, content: heroCard.content, preview: heroCard };

        return {
            composeExtension: {
                type: 'result', 
                attachmentLayout: 'list',
                attachments: [
                    attachment
                ]
            }
        };
    }
}


module.exports.BotActivityHandler = BotActivityHandler;


import { BotFrameworkAdapter, MessageFactory, ConversationParameters, Activity, CardFactory, TurnContext, ConversationReference } from "botbuilder";
import * as debug from "debug";
import * as ACData from "adaptivecards-templating";
import { IConvReference } from "./tableService";

// Import required cards
// tslint:disable-next-line:no-var-requires
const askForHelpCard = require("./cards/askForHelpCard.json");

// tslint:disable-next-line:no-var-requires
require("dotenv").config();

// Initialize debug logging module
const log = debug("msteams");

// tslint:disable-next-line:no-var-requires
const BotConnector = require("botframework-connector");

let adapter: BotFrameworkAdapter;
let connectorClient;


const initConnectorClient = () => {

        adapter = new BotFrameworkAdapter({
            appId: process.env.MICROSOFT_APP_ID,
            appPassword: process.env.MICROSOFT_APP_PASSWORD
        });

        BotConnector.MicrosoftAppCredentials.trustServiceUrl(
            process.env.SERVICE_URL
        );

        connectorClient = adapter.createConnectorClient(process.env.SERVICE_URL as string);

        log("connector client initialized");
};


const sendUserMessage = async (messageTxt: string, userId: string) => {

        const message = MessageFactory.text(messageTxt) as Activity;

        // User Scope
        const conversationParameters = {
            isGroup: false,
            channelData: {
                tenant: {
                    id: process.env.TENANT_ID
                }
            },
            bot: {
                id: process.env.BOT_ID,
                name: process.env.BOT_NAME
            },
            members: [
                {
                    id: userId
                }
            ]
        };

        const parametersTalk = conversationParameters as ConversationParameters;
        const response = await connectorClient.conversations.createConversation(parametersTalk);
        await connectorClient.conversations.sendToConversation(response.id, message);

        log("user message sent");
};


const sendChannelMessage = async (messageTxt: string) => {

        let message = MessageFactory.text(messageTxt) as Activity;

        // Channel Scope
        const conversationParameters = {
            isGroup: true,
            channelData: {
                channel: {
                    id: process.env.TEAMS_CHANNEL_ID
                }
            },
            activity: message
        };

        const conversationParametersReference = conversationParameters as ConversationParameters;
        const response = await connectorClient.conversations.createConversation(conversationParametersReference);

        // Send reply to channel
        message = MessageFactory.text("This is the first reply in a new channel message") as Activity;
        await connectorClient.conversations.sendToConversation(response.id, message);

        log("channel message sent");
};

const notifyChannel = async (nome: string, duvida: string) => {

    // Build Proactive Notify Team Card
    const template = new ACData.Template(askForHelpCard);
    const askForHelpCardExpanded = await template.expand({
    $root: {
        txtnome: nome,
        txtduvida: duvida
    }
    });

    // Reply Back to User
    const AskForHelpAdaptiveCard = CardFactory.adaptiveCard(askForHelpCardExpanded);

    const message = MessageFactory.attachment(AskForHelpAdaptiveCard) as Activity;

    // Channel Scope
    const conversationParameters = {
        isGroup: true,
        channelData: {
            channel: {
                id: process.env.TEAMS_CHANNEL_ID
            }
        },
        activity: message
    } as ConversationParameters;

    const response = await connectorClient.conversations.createConversation(conversationParameters);

    log("channel message sent");

    return response.id;
};

const routeToChannel = async (message: string, conversationid: string) => {

    // Send reply to channel
    const mensagem = MessageFactory.text(message) as Activity;

    await connectorClient.conversations.sendToConversation(conversationid, mensagem);
    log("channel reply sent");

};

const routeToUser = async (conversationid: any, message: string) => {

    const conversationReferenceWeb = JSON.parse(conversationid) as ConversationReference;

    await adapter.continueConversation(conversationReferenceWeb, async turnContext => {
        await turnContext.sendActivity(message);
    });

};

const notifyMeeting = async (reference: IConvReference, card: JSON) => {

    // Build Card
    const notifyCard = CardFactory.adaptiveCard(card);
    const mensagem = MessageFactory.attachment(notifyCard) as Activity;

    // Notify Channel
    await connectorClient.conversations.sendToConversation(reference.channelRef, mensagem);
    log("meeting send to Channel");

    // Notify User
    const conversationReferenceWeb = JSON.parse(reference.convRef) as ConversationReference;

    await adapter.continueConversation(conversationReferenceWeb, async turnContext => {
        await turnContext.sendActivity(mensagem);
    });

};

export {
    sendUserMessage,
    sendChannelMessage,
    initConnectorClient,
    notifyChannel,
    routeToChannel,
    routeToUser,
    notifyMeeting
};


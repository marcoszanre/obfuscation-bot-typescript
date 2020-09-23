import { BotFrameworkAdapter, MessageFactory, ConversationParameters, Activity, CardFactory, ConversationReference } from "botbuilder";
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

// Initialize reference variables for proactive messaging
let adapter: BotFrameworkAdapter;
let connectorClient;

// Initialization function to create the adapter and connector to proactively send messages
const initConnectorClient = () => {

        // Initialize adapter
        adapter = new BotFrameworkAdapter({
            appId: process.env.MICROSOFT_APP_ID,
            appPassword: process.env.MICROSOFT_APP_PASSWORD
        });

        // Trust the service url
        BotConnector.MicrosoftAppCredentials.trustServiceUrl(
            process.env.SERVICE_URL
        );

        // Initialize connector
        connectorClient = adapter.createConnectorClient(process.env.SERVICE_URL as string);

        // Notify success
        log("connector client initialized");

};

// Function to proactively send a new card to the experts channel
const notifyChannel = async (nome: string, duvida: string) => {

    // Build Proactive Notify Team Card and fill it out using adaptive cards templating
    const template = new ACData.Template(askForHelpCard);
    const askForHelpCardExpanded = await template.expand({
    $root: {
        txtnome: nome,
        txtduvida: duvida
    }
    });

    // Create the card
    const AskForHelpAdaptiveCard = CardFactory.adaptiveCard(askForHelpCardExpanded);

    // Attach the card to the message as an Activity
    const message = MessageFactory.attachment(AskForHelpAdaptiveCard) as Activity;

    // Fill out Channel reference variables
    const conversationParameters = {
        isGroup: true,
        channelData: {
            channel: {
                id: process.env.TEAMS_CHANNEL_ID
            }
        },
        activity: message
    } as ConversationParameters;

    // Create a new thread in channel
    const response = await connectorClient.conversations.createConversation(conversationParameters);

    // Notify local process of success
    log("channel message sent");

    // Return the id of the new thread to reference in database for future proactive messages
    return response.id;

};


// Function to route the public user message received from the dialog to the experts channel
const routeToChannel = async (message: string, conversationid: string) => {

    // Create message as an Activity
    const mensagem = MessageFactory.text(message) as Activity;

    // Send proactive message in existing thread in Teams experts channel
    await connectorClient.conversations.sendToConversation(conversationid, mensagem);

    // Notify local process of success
    log("channel reply sent");

};


// Function to route the Teams expert message to public user
const routeToUser = async (conversationid: any, message: string) => {

    // Parse the stringified conversation reference from database
    const conversationReferenceWeb = JSON.parse(conversationid) as ConversationReference;

    // Leverage the adapter to continue the conversation
    await adapter.continueConversation(conversationReferenceWeb, async turnContext => {

        // Send the activity to the public user
        await turnContext.sendActivity(message);
    });

    // Notify local process of success
    log("user message sent");

};


// Function to let participants know that a Teams meeting has been created and share the meeting card
const notifyMeeting = async (reference: IConvReference, card: JSON) => {

    // Build Meeting notify card based on JSON received from function variables
    const notifyCard = CardFactory.adaptiveCard(card);

    // Attach the card to the message as an Activity
    const mensagem = MessageFactory.attachment(notifyCard) as Activity;

    // Notify Teams Experts channel with Meeting Card
    await connectorClient.conversations.sendToConversation(reference.channelRef, mensagem);

    // Notify local process of success
    log("meeting sent to Channel");

    // Notify public User with Meeting Card
    const conversationReferenceWeb = JSON.parse(reference.convRef) as ConversationReference;

    // Leverage the adapter to continue the conversation
    await adapter.continueConversation(conversationReferenceWeb, async turnContext => {

        // Send the activity to the public user
        await turnContext.sendActivity(mensagem);

    });

    // Notify local process of success
    log("meeting sent to user");

};


// Export proactive message functions to be used as a service
export {
    initConnectorClient,
    notifyChannel,
    routeToChannel,
    routeToUser,
    notifyMeeting
};


import { BotDeclaration } from "express-msteams-host";
import { TurnContext, MemoryStorage, TeamsActivityHandler, MessageFactory, ConversationParameters, Activity, BotFrameworkAdapter, TeamsInfo, CardFactory, ConversationReference, ConversationState, StatePropertyAccessor, ActivityTypes } from "botbuilder";
import { initTableSvc, insertConvReference, getConvReference, IConvReference } from "../tableService";
import { DialogSet, DialogState, DialogTurnStatus } from "botbuilder-dialogs";
import { sendChannelMessage, sendUserMessage, initConnectorClient, notifyChannel, routeToChannel, routeToUser, notifyMeeting } from "../connectorClient";
import { generateAnswer } from "../qnaService";
import * as ACData from "adaptivecards-templating";
import webToExpertDialog from "../dialogs/webToExpertDialog";
import { initGraphSvc, createMeeting } from "../graphService";

// Import required cards
const qnaReplyCard = require(".././cards/qnareplyCard.json");
const cancelarExpertCard = require(".././cards/cancelExpertCard.json");
const notifyMeetingsCard = require(".././cards/notifyMeetingCard.json");


/**
 * Implementation for proactive messages Bot
 */
@BotDeclaration(
    "/api/messages",
    new MemoryStorage(),
    process.env.MICROSOFT_APP_ID,
    process.env.MICROSOFT_APP_PASSWORD)

export class ProactiveMessagesBot extends TeamsActivityHandler {
    private readonly conversationState: ConversationState;
    private readonly dialogs: DialogSet;
    private dialogState: StatePropertyAccessor<DialogState>;  

    /**
     * The constructor
     * @param conversationState
     */
    public constructor(conversationState: ConversationState) {
        super();

        this.conversationState = conversationState;
        this.dialogState = conversationState.createProperty("dialogState");
        this.dialogs = new DialogSet(this.dialogState);
        this.dialogs.add(new webToExpertDialog("chat"));

        // Init table service
        initTableSvc();

        // Init Connector Client
        initConnectorClient();

        // Init Graph Client
        initGraphSvc();

        // Set up the Activity processing
        this.onMessage(async (context: TurnContext): Promise<void> => {

            if(context.activity.conversation.conversationType) {

                // Message received in Teams

                // Confirm if message is a card reply or not
                if (context.activity.value) {

                    // Notify conversation to join meeting
                    await getConvReference(context.activity.conversation.id).then(async (ref: IConvReference) => {

                        const randID = Math.floor(Math.random() * 10000 + 1)
                   
                        // Create Graph Meeting
                        const myMeeting = await createMeeting(`Atendimento Contoso ${randID}`);
                        // console.log(myMeeting);

                        // Create Notify Card
                        let template = new ACData.Template(notifyMeetingsCard);
                        let notifyCardExpanded = template.expand({
                        $root: {
                            url: myMeeting
                        }
                        });

                        notifyMeeting(ref, notifyCardExpanded);
                            
                    });

                 } else {

                    const text = await TurnContext.removeRecipientMention(context.activity);
                
                    await getConvReference(context.activity.conversation.id).then(async (ref: IConvReference) => {

                    // console.log(ref.convRef);

                    routeToUser(ref.convRef, text);

                    // console.log(conversationReferenceWeb);

                    // await context.adapter.continueConversation(conversationReferenceWeb, async turnContext => {
                    //     await turnContext.sendActivity("asdasdad");
                    // });
                    
                     });

                }

            } else {
                const dc = await this.dialogs.createContext(context);

                // Confirm if message is a card reply or not
                if (context.activity.value) {

                    // Send Typing Activity
                    await context.sendActivity({type:  ActivityTypes.Typing});

                    await context.sendActivity(`Obrigado ${context.activity.value.txtNome}! Vou escalar agora! Um momento que já te aviso quando um especialista se conectar 😉`);

                    // Send Typing Activity
                    await context.sendActivity({type:  ActivityTypes.Typing});

                    // Prepare storage reference
                    const channelConversationReference = await notifyChannel(context.activity.value.txtNome, context.activity.value.txtDuvida);
                    const webConversationReference = JSON.stringify(TurnContext.getConversationReference(context.activity));
                    const webConversationID = context.activity.conversation.id;

                    insertConvReference(
                        channelConversationReference,
                        webConversationReference,
                        webConversationID
                    )

                    const cancelarCard = CardFactory.adaptiveCard(cancelarExpertCard);
                    await context.sendActivity( { attachments: [cancelarCard] });

                    // routeToChannel("this is a reply", channelConversationReference)

                    // await context.sendActivity("notifyChannel Done 😎");

                    // const dc = await this.dialogs.createContext(context);
                    await dc.beginDialog("chat");

                } else {

                    // const dc = await this.dialogs.createContext(context);

                    // Cancel Dialog execution if user sends cancelar
                    // if (context.activity.text.startsWith("cancelar")) {
                    //     await dc.cancelAllDialogs();
                    // }

                    const results = await dc.continueDialog();

                    // If there's no dialog running, run this
                    if (results.status === DialogTurnStatus.empty) {

                        // Send Typing Activity
                        await context.sendActivity({type:  ActivityTypes.Typing});

                        // Retrieve answer from QnA
                        const answer = await generateAnswer(context);

                        // Build QnA Reply Card
                        let template = new ACData.Template(qnaReplyCard);
                        let qnaReplyCardExpanded = template.expand({
                        $root: {
                            resposta: answer
                        }
                        });
                        
                        // Reply Back to User
                        const QnAReplyAdaptiveCard = CardFactory.adaptiveCard(qnaReplyCardExpanded);
                        await context.sendActivity({ attachments: [QnAReplyAdaptiveCard] });

                    }

                    // await dc.beginDialog("chat"); 

                    // const convReference = JSON.stringify(TurnContext.getConversationReference(context.activity));
                    // console.log(convReference);
                    
                    // routeToUser(convReference, "blablabla");

                }


            }

            /*
            await sendChannelMessage(`Proactive message: ${context.activity.text}`);
            await context.sendActivity("sendChannel Done 😎");

            if (context.activity.conversation.conversationType === "personal") {
                switch (context.activity.text) {
                    case 'get':
                        await this.createUserConversation(context);
                        break;
                    case 'channel':
                        await this.teamsCreateConversation(context);
                        break;
                    case 'senduser':
                        await sendUserMessage(`Proactive message: ${context.activity.text}`, "29:1r48gyAgyrbiAeDNnSVcd99hKNML6XcwBorYH4OOxZjzBCFYHtRKZMW3c2at7SLedQCCYvGTYWbvbw8VT5fBAjA");
                        await context.sendActivity("sendUser Done 😎");
                        break;
                    case 'sendchannel':
                        await sendChannelMessage(`Proactive message: ${context.activity.text}`);
                        await context.sendActivity("sendChannel Done 😎");
                        break;
                    default:
                        insertUserReference(context);
                        break;
                }
            } else {
                // Channel conversation
                let text = TurnContext.removeRecipientMention(context.activity);
                text = text.toLowerCase();
                text = text.trim();

                if (text === "users") {

                    // Store user ids
                    const teamMembers = await TeamsInfo.getTeamMembers(context);
                    teamMembers.forEach(teamMember => {
                        insertUserID(teamMember.aadObjectId as string, teamMember.name, teamMember.id);
                    });
                }
            }
            */

            // await context.sendActivity("thanks for your message 😀");
            // Save state changes
            return this.conversationState.saveChanges(context);
        });

        this.onConversationUpdate(async (context: TurnContext): Promise<void> => {
            if (context.activity.membersAdded && context.activity.membersAdded.length !== 0) {
                for (const idx in context.activity.membersAdded) {
                    if (context.activity.membersAdded[idx].id === context.activity.recipient.id) {
                        await context.sendActivity("Bem-vindo! Fique à vontade para tirar suas dúvidas conosco! 😀");
                    }
                }
            }
        });

        this.onMessageReaction(async (context: TurnContext): Promise<void> => {
            const added = context.activity.reactionsAdded;
            if (added && added[0]) {
                await context.sendActivity({
                    textFormat: "xml",
                    text: `That was an interesting reaction (<b>${added[0].type}</b>)`
                });
            }
        });;
   }

   async teamsCreateConversation(context: TurnContext) {
    // Create Channel conversation
    let message = MessageFactory.text("This is the first channel message") as Activity;
   
    const conversationParameters: ConversationParameters = {
        isGroup: true,
        channelData: {
            channel: {
                id: process.env.TEAMS_CHANNEL_ID
            }
        },
        bot: {
            id: context.activity.recipient.id,
            name: context.activity.recipient.name
        },
        activity: message
    };

    const notifyAdapter = context.adapter as BotFrameworkAdapter;

    const connectorClient = notifyAdapter.createConnectorClient(context.activity.serviceUrl);
    const response = await connectorClient.conversations.createConversation(conversationParameters);
    await context.sendActivity("conversation sent to channel");

    // Send reply to channel
    message = MessageFactory.text("This is the second channel message") as Activity;
    await connectorClient.conversations.sendToConversation(response.id, message);
    
    }

    async createUserConversation(context: TurnContext) {
        // Create User Conversation
        const message = MessageFactory.text("This is a proactive message") as Activity;
       
        const conversationParameters = {
            isGroup: false,
            channelData: {
                tenant: {
                    id: process.env.TENANT_ID
                }
            },
            members: [
                {
                    id: context.activity.from.id,
                    name: context.activity.from.name
                }
            ]
        };
    
        const notifyAdapter = context.adapter as BotFrameworkAdapter;
        const parametersTalk = conversationParameters as ConversationParameters;
    
        const connectorClient = notifyAdapter.createConnectorClient(context.activity.serviceUrl);
        const response = await connectorClient.conversations.createConversation(parametersTalk);
        await connectorClient.conversations.sendToConversation(response.id, message);
        await context.sendActivity("conversation sent to user");
        
        }

}

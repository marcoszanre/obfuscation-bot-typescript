import { BotDeclaration } from "express-msteams-host";
import { TurnContext, MemoryStorage, TeamsActivityHandler, MessageFactory, ConversationParameters, Activity, BotFrameworkAdapter, CardFactory, ConversationState, StatePropertyAccessor, ActivityTypes } from "botbuilder";
import { initTableSvc, insertConvReference, getConvReference, IConvReference } from "../tableService";
import { DialogSet, DialogState, DialogTurnStatus } from "botbuilder-dialogs";
import { initConnectorClient, notifyChannel, routeToUser, notifyMeeting } from "../connectorClient";
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

        // Init dialog parameters and register the experts chat dialog
        this.conversationState = conversationState;
        this.dialogState = conversationState.createProperty("dialogState");
        this.dialogs = new DialogSet(this.dialogState);
        this.dialogs.add(new webToExpertDialog("chat"));

        // Init table service
        initTableSvc();

        // Init Connector Client
        initConnectorClient()

        // Init Graph Client
        initGraphSvc();

        // Set up the Activity processing
        this.onMessage(async (context: TurnContext): Promise<void> => {

            // Public webchat doesn't send a conversationType data, so we're using that to separate Teams chat from Public chat
            if(context.activity.conversation.conversationType) {

                // Message received in Teams
                
                // Confirm if message is a card reply or not
                if (context.activity.value) {

                    // Card submitted, create meeting

                    // Notify conversation to join meeting
                    await getConvReference(context.activity.conversation.id).then(async (ref: IConvReference) => {

                        // Generate Random ID to create meeting name
                        const randID = Math.floor(Math.random() * 10000 + 1)
                   
                        // Create Graph Meeting
                        const myMeeting = await createMeeting(`Atendimento Contoso ${randID}`);

                        // Create Notify Card and fill it out using adaptive cards templating
                        let template = new ACData.Template(notifyMeetingsCard);
                        let notifyCardExpanded = template.expand({
                        $root: {
                            url: myMeeting
                        }
                        });

                        // Notify participants in both thread with meeting join card
                        notifyMeeting(ref, notifyCardExpanded);
                            
                    });

                 } else {

                    // Message received by user, route to appropriate channel

                    // Remove at mention from text
                    const text = await TurnContext.removeRecipientMention(context.activity);
                
                    // Retrieve public user conversation reference
                    await getConvReference(context.activity.conversation.id).then(async (ref: IConvReference) => {

                        // Route message to user based on database conversation reference
                        routeToUser(ref.convRef, text);
                    
                     });

                }

            } else {

                // Create dialog context to be able to continue user conversation, if applicable
                const dc = await this.dialogs.createContext(context);

                // Confirm if message is a card reply or not
                if (context.activity.value) {

                    // Help card submitted by public user
                    // Send Typing Activity
                    await context.sendActivity({type:  ActivityTypes.Typing});

                    // Send acknowledgement message
                    await context.sendActivity(`Obrigado ${context.activity.value.txtNome}! Vou escalar agora! Um momento que já te aviso quando um especialista se conectar 😉`);

                    // Send Typing Activity
                    await context.sendActivity({type:  ActivityTypes.Typing});

                    // Prepare storage reference after creating a new thread in experts channel
                    const channelConversationReference = await notifyChannel(context.activity.value.txtNome, context.activity.value.txtDuvida);
                    const webConversationReference = JSON.stringify(TurnContext.getConversationReference(context.activity));
                    const webConversationID = context.activity.conversation.id;

                    // Insert conversation reference in database
                    insertConvReference(
                        channelConversationReference,
                        webConversationReference,
                        webConversationID
                    )

                    // Prepare cancel card to let user know how to end specialist chat and send to user
                    const cancelarCard = CardFactory.adaptiveCard(cancelarExpertCard);
                    await context.sendActivity( { attachments: [cancelarCard] });

                    // Route existing public user into the dialog with Teams experts
                    await dc.beginDialog("chat");

                } else {

                    // Default text message submitted by public user

                    // If there's a dialog conversation in place, continue it
                    const results = await dc.continueDialog();

                    // If there's no dialog running, run this
                    if (results.status === DialogTurnStatus.empty) {

                        // Send Typing Activity
                        await context.sendActivity({type:  ActivityTypes.Typing});

                        // Retrieve answer from QnA Service
                        const answer = await generateAnswer(context);

                        // Build QnA Reply Card and fill it out using adaptive cards templating
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
                }
            }

            // Save state changes for dialog processing
            return this.conversationState.saveChanges(context);
            
        });

        // Process conversation update when public user initiates chat
        this.onConversationUpdate(async (context: TurnContext): Promise<void> => {
            if (context.activity.membersAdded && context.activity.membersAdded.length !== 0) {
                for (const idx in context.activity.membersAdded) {
                    if (context.activity.membersAdded[idx].id === context.activity.recipient.id) {

                        // Send to the public user an introduction message
                        await context.sendActivity("Bem-vindo! Fique à vontade para tirar suas dúvidas conosco! 😀");
                    }
                }
            }
        });

        // Default reaction received by bot generated by the Teams generator, but this doesn't apply to public chat
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
}

import { DialogTurnResult, ComponentDialog, TextPrompt, WaterfallDialog, WaterfallStepContext } from "botbuilder-dialogs";
import { routeToChannel } from "../connectorClient";
import { getConvChannelReference, IConvReference } from "../tableService";

// Init dialog ids
const TEXT_PROMPT = 'TEXT_PROMPT';
const WATERFALL_DIALOG = 'WATERFALL_DIALOG';

export default class webToExpertDialog extends ComponentDialog {
    constructor(dialogId: string) {
        super(dialogId);

        // Create a new waterfall dialog, add steps and register it as the way of processing it
        this.addDialog(new TextPrompt(TEXT_PROMPT));
        this.addDialog(new WaterfallDialog(WATERFALL_DIALOG, [
            this.requestStep.bind(this),
            this.processStep.bind(this)
        ]));

        // Make waterfall the initial dialog
        this.initialDialogId = WATERFALL_DIALOG;

    }

    // Request step is a blank prompt sent to user to be able to collect his input, so that it looks like a conversation
    async requestStep(stepContext: WaterfallStepContext): Promise<DialogTurnResult> {

        // Blank input sent and awaiting user response
        return await stepContext.prompt(TEXT_PROMPT, "");

    }

    // Process step will make sure that user is not interested in terminating the conversation and will route messages to appropriate channels
    async processStep(stepContext: WaterfallStepContext): Promise<DialogTurnResult> {

        // Confirm that the user has not sent the cancellation text, as advised by the card
        if (stepContext.result == "cancelar") {

            // Public user wants to terminate the conversation, retrieve reference from database to let everyone know of it and end this dialog
            await getConvChannelReference(stepContext.context.activity.conversation.id).then(async (ref: IConvReference) => {

                // Teams channel reference retrieved and user message is routed to the specialist
                routeToChannel("Esta conversa foi encerrada pelo usuário, obrigado!", ref.channelRef)
                
            });

            // Let public user know that this conversation has been terminated
            await stepContext.context.sendActivity("Esta conversa foi encerrada. Obrigado! ✔")
            
            // Terminate dialog, so that if a user sends message again, the QnA service will be the route
            return await stepContext.endDialog();

        } else {

            // Public user wants to continue conversation with expert, route the message to the channel and reprompt user with a blank prompt for new messages

            // Retrieve reference from database to route message to Teams experts
            await getConvChannelReference(stepContext.context.activity.conversation.id).then(async (ref: IConvReference) => {

                // Teams channel reference retrieved and user message is routed to the specialist
                routeToChannel(stepContext.result, ref.channelRef)
                
            });
            
            // Replace dialog to restart it and continue to reprompt message to the user
            return await stepContext.replaceDialog(WATERFALL_DIALOG);

        }
    }
}

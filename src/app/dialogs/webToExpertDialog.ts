import { Dialog, DialogContext, DialogTurnResult, ComponentDialog, TextPrompt, WaterfallDialog, WaterfallStepContext } from "botbuilder-dialogs";
import { ConsoleTranscriptLogger } from "botbuilder";
import { routeToChannel, routeToUser } from "../connectorClient";
import { getConvChannelReference, IConvReference } from "../tableService";

const TEXT_PROMPT = 'TEXT_PROMPT';
const WATERFALL_DIALOG = 'WATERFALL_DIALOG';

export default class webToExpertDialog extends ComponentDialog {
    constructor(dialogId: string) {
        super(dialogId);

        this.addDialog(new TextPrompt(TEXT_PROMPT));
        this.addDialog(new WaterfallDialog(WATERFALL_DIALOG, [
            this.requestStep.bind(this),
            this.processStep.bind(this)
        ]));

        this.initialDialogId = WATERFALL_DIALOG;

    }

    async requestStep(stepContext: WaterfallStepContext): Promise<DialogTurnResult> {
        // return await stepContext.prompt(TEXT_PROMPT, "digite sua próxima mensagem");
        return await stepContext.prompt(TEXT_PROMPT, "");
    }

    async processStep(stepContext: WaterfallStepContext): Promise<DialogTurnResult> {
        // console.log(stepContext.result);

        if (stepContext.result == "cancelar") {

            await getConvChannelReference(stepContext.context.activity.conversation.id).then(async (ref: IConvReference) => {

                routeToChannel("Esta conversa foi encerrada pelo usuário, obrigado!", ref.channelRef)
                
            });

            await stepContext.context.sendActivity("Esta conversa foi encerrada. Obrigado! ✔")
            
            return await stepContext.endDialog();

        } else {

            // console.log(stepContext.context.activity.conversation.id);

            await getConvChannelReference(stepContext.context.activity.conversation.id).then(async (ref: IConvReference) => {

                routeToChannel(stepContext.result, ref.channelRef)
                
            });
            
            return await stepContext.replaceDialog(WATERFALL_DIALOG);

        }

    }

}

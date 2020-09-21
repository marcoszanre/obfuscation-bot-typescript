
// tslint:disable-next-line:no-var-requires
const { QnAMaker } = require("botbuilder-ai");

import * as debug from "debug";
import { TurnContext } from "botbuilder";

const log = debug("msteams");
let answer: string = "";

const qnaMaker = new QnAMaker({
    knowledgeBaseId: process.env.QNA_KNOWLEDGE_BASE_ID,
    endpointKey: process.env.QNA_ENDPOINT_KEY,
    host: process.env.QNA_ENDPOINT_HOSTNAME
});

const generateAnswer = async (context: TurnContext) => {
    log("generate answer called");

    // const qnaResults = await qnaMaker.getAnswers(question);
    const qnaResults = await qnaMaker.getAnswers(context);

    // If answers were found
    if (qnaResults[0]) {
        answer = qnaResults[0].answer;

    // If no answers were returned from QnA Maker, reply with help.
    } else {
        answer = "Desculpe, não encontrei nenhuma resposta";
    }

    return answer;
};

export { generateAnswer };

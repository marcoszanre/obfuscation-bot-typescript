// tslint:disable-next-line:no-var-requires
const { QnAMaker } = require("botbuilder-ai");

import * as debug from "debug";
import { TurnContext } from "botbuilder";

const log = debug("msteams");

// Initialize answer variable to be filled out by QnA Service Call
let answer: string = "";


// Initialize qna maker instance
const qnaMaker = new QnAMaker({
    knowledgeBaseId: process.env.QNA_KNOWLEDGE_BASE_ID,
    endpointKey: process.env.QNA_ENDPOINT_KEY,
    host: process.env.QNA_ENDPOINT_HOSTNAME
});


// Function to generate a new qna service answer
const generateAnswer = async (context: TurnContext) => {

    // Notify that the qna service has been called
    log("generate answer called");

    // Store qna maker results
    const qnaResults = await qnaMaker.getAnswers(context);

    // If answers were found, return the most relevant answer
    if (qnaResults[0]) {

        // Update answer variable with the most relevant answer
        answer = qnaResults[0].answer;

    // If no answers were returned from QnA Maker, reply with help.
    } else {

        // Update answer variable with the default answer
        answer = "Desculpe, não encontrei nenhuma resposta";
    }

    // Return answer to be used in the application
    return answer;
};


// Export functions
export { generateAnswer };

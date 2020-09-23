import * as debug from "debug";

// tslint:disable-next-line:no-var-requires
const azure = require("azure-storage");

// Initialize debug logging module
const log = debug("msteams");

// Initialize Table Service
const tableSvc = azure.createTableService(process.env.STORAGE_ACCOUNT_NAME, process.env.STORAGE_ACCOUNT_ACCESSKEY);

// Function to initialize the table service and return success to the application
const initTableSvc = () => {
    tableSvc.createTableIfNotExists("proactiveTable", (error) => {
        if (!error) {

          // Table exists or created
          log("table service done");
        }
    });

};


// Function to insert a public user and experts team conversation in database, to be used later for proactive messaging
const insertConvReference = (channelConversationReference: string, webConversationReference: string, webConversationID: string) => {

        const convReference = {
            PartitionKey: {_: "convReference"},
            RowKey: {_: webConversationID},
            channelRef: {_: channelConversationReference},
            webRef: {_: webConversationReference},
        };

        tableSvc.insertEntity("proactiveTable", convReference, (error) => {
            if (!error) {
              // Entity inserted
              log("success!");
            } else {
                log(error);
            }
        });

};


// Function to retrieve the conversation reference by the teams experts, based on the Teams channel thread id
const getConvReference = async (channelRef: string) => {

    return new Promise((resolve) => {

        const query = new azure.TableQuery()
            .where("channelRef eq ?", channelRef);

        tableSvc.queryEntities("proactiveTable", query, null, (error, result) => {
            if (!error) {
              // query was successful
              const convReturnReference: IConvReference = {
                convRef: result.entries[0].webRef._,
                channelRef: result.entries[0].channelRef._
              };
              resolve(convReturnReference);
            }
        });
    });
};


// Conversation Reference interface
interface IConvReference {
    convRef: string;
    channelRef: string;
}


// Function to retrieve the conversation reference by the public user, based on the public chat bot conversation id to post to the Teams channel
const getConvChannelReference = async (webconvid: string) => {

    return new Promise((resolve) => {

        tableSvc.retrieveEntity("proactiveTable", "convReference", webconvid, (error, result) => {
            if (!error) {
              // query was successful
              const convReturnReference: IConvReference = {
                convRef: result.webRef._,
                channelRef: result.channelRef._
              };
              resolve(convReturnReference);
            }
        });
    });

};


// Export table service functions
export {
    initTableSvc,
    insertConvReference,
    getConvReference,
    IConvReference,
    getConvChannelReference
};

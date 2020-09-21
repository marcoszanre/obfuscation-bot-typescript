import { TurnContext } from "botbuilder";
import * as debug from "debug";

// tslint:disable-next-line:no-var-requires
const azure = require("azure-storage");

// Initialize debug logging module
const log = debug("msteams");

const tableSvc = azure.createTableService(process.env.STORAGE_ACCOUNT_NAME, process.env.STORAGE_ACCOUNT_ACCESSKEY);

const initTableSvc = () => {
    tableSvc.createTableIfNotExists("proactiveTable", (error, result, response) => {
        if (!error) {
          // Table exists or created
          log("table service done");
        }
    });
};

const insertConvReference = (channelConversationReference: string, webConversationReference: string, webConversationID: string) => {

        const convReference = {
            PartitionKey: {_: "convReference"},
            RowKey: {_: webConversationID},
            channelRef: {_: channelConversationReference},
            webRef: {_: webConversationReference},
        };

        tableSvc.insertEntity("proactiveTable", convReference, (error, result, response) => {
            if (!error) {
              // Entity inserted
              log("success!");
            } else {
                log(error);
            }
        });
};


const getConvReference = async (channelRef: string) => {

    return new Promise((resolve, reject) => {

        const query = new azure.TableQuery()
            .where("channelRef eq ?", channelRef);

        tableSvc.queryEntities("proactiveTable", query, null, (error, result, response) => {
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

interface IConvReference {
    convRef: string;
    channelRef: string;
}

const getConvChannelReference = async (webconvid: string) => {

    return new Promise((resolve, reject) => {

        tableSvc.retrieveEntity("proactiveTable", "convReference", webconvid, (error, result, response) => {
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


export {
    initTableSvc,
    insertConvReference,
    getConvReference,
    IConvReference,
    getConvChannelReference
};

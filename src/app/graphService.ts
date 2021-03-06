import * as debug from "debug";

// tslint:disable-next-line:no-var-requires
const axios = require("axios");

// tslint:disable-next-line:no-var-requires
require("dotenv").config();

// tslint:disable-next-line:no-var-requires
const qs = require("qs");

// Initialize debug logging module
const log = debug("msteams");

// Create access token variable to store Graph permissions
let accessToken: string;


// Initialize Graph Service and retrieve a token to be used to provision meetings
const initGraphSvc = async () => {

    // Retrieve token response from directory token service and update the accessToken variable with it
    const tokenEndpointResponse = await callTokenEndpoint();
    accessToken = tokenEndpointResponse.data.access_token;

    // Notify success
    log("Graph Service initialized");
};


// Function to retrieve access token leveraging a service account and an application registration with appropriate createOnlineMeeting permissions
const callTokenEndpoint = async () => {
    try {
      return await axios({
        method: "post",
        url: `https://login.microsoftonline.com/${process.env.TENANT_ID}/oauth2/v2.0/token`,
        data: qs.stringify({
            grant_type: "password",
            client_id: process.env.GRAPH_CLIENT_ID,
            client_secret: process.env.GRAPH_CLIENT_SECRET,
            username: process.env.GRAPH_USERNAME,
            password: process.env.GRAPH_USERPASSWORD,
            scope: "https://graph.microsoft.com/.default"
        }),
        headers: {
          "content-type": "application/x-www-form-urlencoded;charset=utf-8"
        }
    });
    } catch (error) {
      log(error);
    }

};


// Function to call the Microsoft Graph endpoint with existing token to provision a new Teams meeting
const createOnlineMeeting = async (token, meetingname) => {
    try {
      return await axios({
        method: "post",
        url: "https://graph.microsoft.com/v1.0/me/onlineMeetings",
        data: ({
            startDateTime: "2020-07-12T14:30:34.2444915-07:00",
            endDateTime: "2020-07-12T15:00:34.2464912-07:00",
            subject: meetingname
        }),
        headers: {
          "content-type": "application/json",
          "Authorization": "Bearer " + token
        }
    });
    } catch (error) {
      log(error);
    }
};


// Create Teams meeting and retrieve the id to be used later
const createMeeting = async (meetingname: string) => {
    const meetingData = await createOnlineMeeting(accessToken, meetingname);
    return meetingData.data.joinUrl;
};


// Export the appropriate functions to be used in the application
export {
    createMeeting,
    initGraphSvc
 };

**Important**: this sample code is provided as is.

# Obfuscation Bot - a Microsoft Teams and Public Chat bot

The obfuscation bot is an example of how a Bot Framework bot is able to interact with two channels (public web chat and a Microsoft Teams team) to enable public users to interact with experts. In addition, this bot also connects with Microsoft Graph, based on a resource account, to dynamically create Teams meetings that both experts and public users can join to interact over audio/video and screen sharing.


## How does it work?

When the bot receives an adaptive card user submission, it will create a new thread in the Teams expert channel and store both conversation references in an Azure Table Storage, to refer to when posting proactive messages between the user and the channel.

In order to provide a proper user experience when chatting with experts, the bot leverages dialogs and prompts (blank prompts actually), so that, whenever a user sends a message in the public bot, or the Teams channel, the bot refers to the conversation reference and sends a proactive message in the other channel with the content it received.

In addition, if an expert decides that the conversation should rather happen over audio/video or even screen sharing, they can, by clicking on the adaptive card, provision a new Teams meeting, that will be sent in both threads, that participants can reference to connect for this particular case.

If a user wants to leave the dialog and return to the QnA experience, they just need to send "cancelar" and the bot will terminate the dialog and let both threads know that the conversation has ended. The Bot registration in Teams also has some "answer suggestions" to facilitate the interaction between the expert and the user. 


## Components

The Obfuscation Bot solution is comprised of the following components:

* A `bot`, built with Bot Framework, installed in a Teams channel and exposed in the public web chat channel
* A `Microsoft Teams team`, where experts, and anyone who will interact with the bot, join to be notified and connect over user conversations
* An `Azure QnA Service`, that hosts the knowledge base and answer user questions
* An `Azure Table Storage`, to store the conversation references for the public chat and the Teams thread.


## Enviroment variables

The following environment variables, including the ones already created by the Teams Yeoman Generator, have been used in this application:

STORAGE ACCOUNT
* `STORAGE_ACCOUNT_NAME=` Azure Storage account name
* `STORAGE_ACCOUNT_ACCESSKEY=` Azure Storage account access key

CONNECTOR CLIENT (TO SEND PROACTIVE NOTIFICATIONS)
* `TEAMS_CHANNEL_ID=` experts team channel id (**after running it through URL decode**)
* `SERVICE_URL=` channel connection reference obtained through the context received by the bot (may vary per tenant)

QNA SERVICE
* `QNA_KNOWLEDGE_BASE_ID=` QnA Service Knowledge Base ID
* `QNA_ENDPOINT_KEY=` QnA Service Endpoint Key
* `QNA_ENDPOINT_HOSTNAME=` QnA Service Hostname

GRAPH SERVICE
* `GRAPH_CLIENT_ID=` Azure AD Application Registration ID with appropriate online meeting creation permissions
* `GRAPH_CLIENT_SECRET=` Azure AD Application Registration secret (the consent was provided by an administrator in advance)
* `GRAPH_USERNAME=` Resource account username (the bot identity Graph endpoint is in Beta as of the publishing of this demo)
* `GRAPH_USERPASSWORD=` Resource account password (doesn't support Multi-Factor Authentication)


## How to run it locally

In order to run this repository locally, follow the steps below:

* Download a copy of this repository
* Create in Azure a Bot Channel Registration and enable two channels (Teams and WebChat) - WebChat should already be enabled
* Create in Azure a Storage Account
* Create in Azure and QnA Maker a QnA Service and a Knowledge Base
* Create a new manifest pointing to your bot and sideload (or install) it in a Microsoft Teams team
* Select the channel in this team where the bot will posts cards, copy its URL and decode it
* Create a new Azure AD Application Registration
* Assign the OnlineMeetings.ReadWrite delegated permission (or application depending on how you plan to obtain the token) and proceed to provide the admin consent to the application (or manual consent, if applicable)
* Create a new resource account and assign a Teams license to it
* Create a .env file and fill out required variables (besides the one listed above, Yeoman Teams generator also creates `HOSTNAME`, `APPLICATION_ID`, `PACKAGE_NAME`, `MICROSOFT_APP_ID`, `MICROSOFT_APP_PASSWORD` and `PORT`), so if you want to create a new Teams bot through the Yeoman generator, copy the env file created and add the variables listed above, that would work
* Run `ngrok` locally exposing the desired port (Yeoman Teams generator uses 3007, but that will be based on your `PORT` environment variable)
* Initiate the bot running `gulp serve`, and make sure that all services (Graph, Table and Connector) have initiated correctly
* Open the `Test in Web Chat` experience or copy the iframe reference of the Web Chat to add it to an existing web page and start interacting with it.


## Expected Conversation Flow

The following conversation flow has been planned through the building process of this bot:

* User receives a bot message when initiating a new conversation
* User sends a question to the bot, that is answered with a card
* If a user wants to escalate the conversation, he/she can escalate in the card by filling out the name and question
* The bot will process the card submission, create a new thread in the chosen Teams team and initiate the user/expert dialog routing
* If a user sends a message to the bot, it will be sent to the appropriate thread, and if an expert mentions the bot, the message will be sent to the user as well
* If an expert clicks on the `Meeting` button in the card of the thread, a new Teams meeting will be created and a card with join coordinates will be sent in both threads, which user and expert can join, if applicable
* User can send `cancelar` to terminate the conversation, and bot will notify in both threads that the conversation has ended
* User can ask questions again to the knowledge base.

## To Do

The following scenarios, though important, have not been implemented yet in this demo:

* Bot validation to check all card fields have been filled out by user
* Graph Service concurrent logic, to handle large loads of provisioning meeting scenarios, if applicable
* Connector client throttling handling, to handle large loads of proactive messages, if applicable
* Graph Service token renewal implementation (e.g. every 45 minutes)
* Table Service no conversation reference found error handling (this shouldn't happen though) 
* Table Service active chat boolean column
* Teams bot checks if a chat is still active, to not try to send webchat user messages in case a thread has already been closed
* Logic to confirm if the user is still active in webchat, or even left the page, to terminate the conversation automatically
* Move bot dialog storage implementation from memory to a persistent storage (e.g. Azure Storage).


## Extension opportunities

Following are ideas, that could be considered to extend the reach of this solution, to target even more scenarios:

* Create and enable a database of pre-configured answers (the bot app registration in Teams help text is limited to 32 characters), that users could type in the bot from Teams, and will be replaced with the configured answer
* Create a code (e.g. {user}) that users could type to have the bot replace the text with the public user name
* Expose to the Teams bot internal services (e.g. check user data against internal information from the chat itself - "@bot, what's userid status?").


## Demo screenshots

1. User interaction with web channel (right) and can escalate questions, which are sent to the experts channel in Teams:

![architecture overview](https://github.com/marcoszanre/obfuscation-bot-typescript/blob/master/demo-1.png/)


2. Experts can mention the bot in Teams and user can send messages to the bot and both will be routed:

![architecture overview](https://github.com/marcoszanre/obfuscation-bot-typescript/blob/master/demo-2.png/)


3. Teams meetings are created and coordinates are made available in both channels:

![architecture overview](https://github.com/marcoszanre/obfuscation-bot-typescript/blob/master/demo-3.png/)


## Architecture overview
![architecture overview](https://github.com/marcoszanre/obfuscation-bot-typescript/blob/master/architecture-overview.png/)


## References

Following are some of the references used in this project:

* [ngrok](https://ngrok.io)
* [Microsoft Teams official documentation](https://developer.microsoft.com/en-us/microsoft-teams)
* [Microsoft Teams Yeoman generator Wiki](https://github.com/PnP/generator-teams/wiki)
* [Create Online Meeting Microsoft Graph Endpoint](https://docs.microsoft.com/en-us/graph/api/application-post-onlinemeetings?view=graph-rest-1.0&tabs=http)
* [Proactive Messages in Microsoft Teams Typescript Bots YouTube Demo](https://www.youtube.com/watch?v=kEL_FUlRpY0&feature=youtu.be)
* [Azure QnA Maker](https://docs.microsoft.com/en-us/azure/cognitive-services/qnamaker/overview/overview)
* [Azure Table Storage](https://docs.microsoft.com/en-us/azure/storage/tables/table-storage-overview)


If you have any questions/suggestions, feel free to share them, **thanks**!
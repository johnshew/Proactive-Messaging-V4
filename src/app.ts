import * as dotenv from 'dotenv';
import * as fs from 'fs';
import * as path from 'path';
import * as restify from 'restify';
import * as util from 'util';

import { ActivityTypes, BotAdapter, BotFrameworkAdapter, ConversationReference, TurnContext } from 'botbuilder';
import { MicrosoftAppCredentials } from 'botframework-connector';

const ENV_FILE = path.join(__dirname, '../.env');
dotenv.config({ path: ENV_FILE });
var appId = process.env.appId;
var appPassword = process.env.appPassword;
if (!appId || !appPassword) { throw new Error('No app credentials.'); process.exit(); }
var botPort = process.env.botport || process.env.BOTPORT || 3978;

let adapter = new BotFrameworkAdapter({ appId: appId, appPassword: appPassword });

class Bot {

    constructor(private adapter: BotAdapter) { }

    async onTurn(turnContext: TurnContext) {
        try {
            if (turnContext.activity.type === ActivityTypes.Message) {
                const utterance = (turnContext.activity.text || '').trim().toLowerCase();
                if (utterance == 'subscribe') {
                    let conversationReference = TurnContext.getConversationReference(turnContext.activity);
                    await util.promisify(fs.writeFile)('conversation-reference.json', JSON.stringify(conversationReference));
                    await turnContext.sendActivity('Subscribed.');
                } else if (utterance == 'notify') {
                    let result = await util.promisify(fs.readFile)('conversation-reference.json');
                    let conversationReference = <ConversationReference>JSON.parse(result.toString());
                    MicrosoftAppCredentials.trustServiceUrl(conversationReference.serviceUrl);
                    await this.adapter.continueConversation(conversationReference, async (turnContext) => {
                        await turnContext.sendActivity('Notified.');
                    });
                } else {
                    await turnContext.sendActivity('Hello.  You can say subscribe or notify.');
                }
            }
        } catch (err) { console.log(`Error: ${err}`) }
    }
}

let bot = new Bot(adapter);

let server = restify.createServer()

server.post('/api/messages', (req, res) => {
    adapter.processActivity(req, res, async (turnContext) => {
        // Route the message to the bot's main handler.
        await bot.onTurn(turnContext);
    });
});

server.listen(botPort, () => {
    console.log(`${server.name} listening to ${server.url}`);
});


// Import required packages
import * as restify from "restify";

// Import required bot services.
// See https://aka.ms/bot-services to learn more about the different parts of a bot.
import {
  CloudAdapter,
  ConfigurationServiceClientCredentialFactory,
  ConfigurationBotFrameworkAuthentication,
  TurnContext,
  MemoryStorage,
  // TeamsSSOTokenExchangeMiddleware
} from "botbuilder";

import { SearchApp } from "./searchApp";
import * as path from "path";
import { config } from "dotenv";
/////////////////////////////////////////////////////////////////

const ENV_FILE = path.join(__dirname, "..", "/env/.env.local.user");
config({ path: ENV_FILE });

const credentialsFactory = new ConfigurationServiceClientCredentialFactory({
  MicrosoftAppId: process.env.BOT_ID,
  MicrosoftAppPassword: process.env.BOT_PASSWORD,
  MicrosoftAppType: "MultiTenant",
});
const botFrameworkAuthentication = new ConfigurationBotFrameworkAuthentication(
  {},
  credentialsFactory
);
const adapter = new CloudAdapter(botFrameworkAuthentication);
const memoryStorage = new MemoryStorage();
// const tokenExchangeMiddleware = new TeamsSSOTokenExchangeMiddleware(memoryStorage, process.env.CONNECTION_NAME);

const onTurnErrorHandler = async (context: TurnContext, error: Error) => {
  console.error(`\n [onTurnError] unhandled error: ${error}`);
  await context.sendTraceActivity(
    "OnTurnError Trace",
    `${error}`,
    "https://www.botframework.com/schemas/error",
    "TurnError"
  );
  await context.sendActivity(`The bot encountered unhandled error:\n ${error.message}`);
  await context.sendActivity("To continue to run this bot, please fix the bot source code.");
};
adapter.onTurnError = onTurnErrorHandler;

const searchApp = new SearchApp();
const server = restify.createServer();
server.use(restify.plugins.bodyParser());
server.listen(process.env.port || process.env.PORT || 3978, () => {
  console.log(`\nBot Started, ${server.name} listening to ${server.url}`);
});
server.post("/api/messages", async (req, res) => {
  await adapter.process(req, res, async (context) => {
    await searchApp.run(context);
  });
});

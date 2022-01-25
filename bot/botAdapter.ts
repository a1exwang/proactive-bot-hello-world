// Import required bot services.
// See https://aka.ms/bot-services to learn more about the different parts of a bot.
import { BotFrameworkAdapter, TurnContext, WebRequest, WebResponse } from "botbuilder";
import { NotificationSender } from "./notificationSender";
import { ConversationReferenceStore } from "./store";

// This bot's main dialog.
import { TeamsBot } from "./teamsBot";

// Create adapter.
// See https://aka.ms/about-bot-adapter to learn more about adapters.
const adapter = new BotFrameworkAdapter({
  appId: process.env.BOT_ID,
  appPassword: process.env.BOT_PASSWORD,
});

// Catch-all for errors.
const onTurnErrorHandler = async (context: TurnContext, error: Error) => {
  // This check writes out errors to console log .vs. app insights.
  // NOTE: In production environment, you should consider logging this to Azure
  //       application insights.
  console.error(`\n [onTurnError] unhandled error: ${error}`);

  // Send a trace activity, which will be displayed in Bot Framework Emulator
  await context.sendTraceActivity(
    "OnTurnError Trace",
    `${error}`,
    "https://www.botframework.com/schemas/error",
    "TurnError"
  );

  // Send a message to the user
  await context.sendActivity(`The bot encountered unhandled error:\n ${error.message}`);
  await context.sendActivity("To continue to run this bot, please fix the bot source code.");
};

// Set the onTurnError for the singleton BotFrameworkAdapter.
adapter.onTurnError = onTurnErrorHandler;

const conversationReferenceStore = new ConversationReferenceStore();
// Create the bot that will handle incoming messages.
const bot = new TeamsBot(conversationReferenceStore);
const notificationSender = new NotificationSender(adapter);

export async function handleBotRequest(req: WebRequest, res: WebResponse): Promise<void> {
  await adapter.processActivity(req, res, async (context) => {
    await bot.run(context);
  });
}
export async function handleNotification(req: WebRequest, res: WebResponse): Promise<void> {
  notificationSender.sendNotification(
    conversationReferenceStore.get(),
    "Hello world!\nYou've received a notification triggered by API."
  );
}

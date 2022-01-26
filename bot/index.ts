// Import required packages
import * as restify from "restify";

// Import required bot services.
// See https://aka.ms/bot-services to learn more about the different parts of a bot.
import { BotFrameworkAdapter } from "botbuilder";

import { ConversationReferenceStore } from "./store";
import { TeamsBot } from "./teamsBot";
import { NotificationSender } from "./notificationSender";

// Create HTTP server.
const server = restify.createServer();
server.listen(process.env.port || process.env.PORT || 3978, () => {
  console.log(`\nBot Started, ${server.name} listening to ${server.url}`);
});

// Create adapter.
// See https://aka.ms/about-bot-adapter to learn more about adapters.
const adapter = new BotFrameworkAdapter({
  appId: process.env.BOT_ID,
  appPassword: process.env.BOT_PASSWORD,
});

// Create conversation reference storage
const conversationReferenceStore = new ConversationReferenceStore();
// Create the bot that will handle incoming messages.
const bot = new TeamsBot(adapter, conversationReferenceStore);
// Create notification sender to proactively send outgoing messages.
const notificationSender = new NotificationSender(adapter);

// Listen for incoming requests.
server.post("/api/messages", async (req, res) => {
  return await adapter.processActivity(req, res, async (context) => {
    await bot.run(context);
  });
});

// Send notification
server.post("/api/notification", async (req, res) => {
  await notificationSender.sendNotification(
    conversationReferenceStore.get(),
    "Hello world!\nYou've received a notification triggered by API."
  );
  res.json({});
});

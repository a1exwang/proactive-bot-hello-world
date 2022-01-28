// Import required packages
import * as restify from "restify";
import * as fs from "fs";

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
const botId = process.env.BOT_ID;
const adapter = new BotFrameworkAdapter({
  appId: botId,
  appPassword: process.env.BOT_PASSWORD,
});

// Create conversation reference storage
const conversationReferenceStore = new ConversationReferenceStore();
// Create the bot that will handle incoming messages.
const bot = new TeamsBot(adapter, conversationReferenceStore, botId);
// Create notification sender to proactively send outgoing messages.
const notificationSender = new NotificationSender(adapter);

// Bot listens for incoming requests.
server.post("/api/messages", async (req, res) => {
  return await adapter.processActivity(req, res, async (context) => {
    await bot.run(context);
  });
});

// HTTP trigger for the notification.
server.post("/api/notification", async (req, res) => {
  const notificationText = "Hello world!\nYou've received a notification triggered by API.";
  const ref1 = JSON.parse(fs.readFileSync("channelConversationRef.json", {"encoding": "utf-8"}));
  await notificationSender.sendNotification(
    // conversationReferenceStore.get(),
    ref1,
    notificationText,
  );
  res.json({});
});

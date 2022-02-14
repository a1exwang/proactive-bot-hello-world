// Import required packages
import * as restify from "restify";

// Import required bot services.
// See https://aka.ms/bot-services to learn more about the different parts of a bot.
import { BotFrameworkAdapter, MessageFactory } from "botbuilder";

import { ConversationReferenceFileStore } from "./sdk/conversationReferenceFileStore";
import { TeamsBot } from "./teamsBot";
import { NotificationSender } from "./sdk/notificationSender";
import { getTeamMemberInfoByEmail } from "./sdk/botUtils";

// Create HTTP server.
const server = restify.createServer();
server.use(restify.plugins.bodyParser({ mapParams: false }));
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
const conversationReferenceStore = new ConversationReferenceFileStore("ref.json");
// Create the bot that will handle incoming messages.
const bot = new TeamsBot(adapter, conversationReferenceStore, process.env.BOT_ID);
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
  const refs = await conversationReferenceStore.list();
  // Developers can also getContext() and then call TeamsInfo APIs with the context to list member and channels.
  for (const ref of refs) {
    const receiverConversationId = await getTeamMemberInfoByEmail(adapter, ref, req.body.receiver);
    if (receiverConversationId) {
      const message = MessageFactory.text(req.body.content);
      await notificationSender.sendNotificationToMember(ref, receiverConversationId, message);
    }
  }

  res.json({});
});


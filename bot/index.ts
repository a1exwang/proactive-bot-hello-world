// Import required packages
import * as restify from "restify";

// Import required bot services.
// See https://aka.ms/bot-services to learn more about the different parts of a bot.
import { BotFrameworkAdapter, ConversationReference, MessageFactory, teamsGetTeamId, TeamsInfo, TurnContext } from "botbuilder";

import { ConversationReferenceFileStore } from "./sdk/conversationReferenceFileStore";
import { TeamsBot } from "./teamsBot";
import { NotificationSender } from "./sdk/notificationSender";
import { getTeamMemberInfoByEmail } from "./sdk/botUtils";
import { SettingsStorage as SettingsStore } from "./sdk/settingsStorage";
import { Settings } from "./sdk/cards";

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
adapter.use(async (context: TurnContext, next) => {
  console.log(JSON.stringify(context.activity, null, 2));
  await next();
});

// Create conversation reference storage
const conversationReferenceStore = new ConversationReferenceFileStore("../ref.json");
const settingsStore = new SettingsStore();
// Create the bot that will handle incoming messages.
const bot = new TeamsBot(adapter, conversationReferenceStore, settingsStore, process.env.BOT_ID);
// Create notification sender to proactively send outgoing messages.
const notificationSender = new NotificationSender(adapter);

// Bot listens for incoming requests.
server.post("/api/messages", async (req, res) => {
  return await adapter.processActivity(req, res, async (context) => {
    await bot.run(context);
  });
});

// HTTP triggers for the notification.
server.post("/api/notify/member", async (req, res) => {
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

server.post("/api/notify/default", async (req, res) => {
  const refs = await conversationReferenceStore.list();
  const settings: Settings = await settingsStore.get();
  for (const ref of refs) {
    // TODO: check or convert
    if (ref.conversation.id in settings.teams) {
      const channels = settings.teams[ref.conversation.id];
      for (const channel of channels) {
        if (channel.subscribed) {
          const newRef: ConversationReference = JSON.parse(JSON.stringify(ref));
          newRef.conversation.id = channel.id;
          await adapter.continueConversation(
            newRef,
            async (context: TurnContext) => {
              const message = MessageFactory.text(req.body.content);
              await context.sendActivity(message);
            }
          );
        }
      }
    }
  }

  res.json({});
});


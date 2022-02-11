import * as restify from "restify";
import { getTeamMemberInfoByEmail } from "./botUtils";
import { BotFrameworkAdapter, MessageFactory } from "botbuilder";
import { TeamsBot } from "./teamsBot";
import { ConversationReferenceFileStore } from "./store";
import { NotificationSender } from "./notificationSender";

const adapter = new BotFrameworkAdapter({
  appId: process.env.BOT_ID,
  appPassword: process.env.BOT_PASSWORD,
});

const store = new ConversationReferenceFileStore("ref.json");
const notificationSender = new NotificationSender(adapter);
const bot = new TeamsBot(adapter, store, process.env.BOT_ID);

// Create HTTP server.
const server = restify.createServer();
server.use(restify.plugins.bodyParser({ mapParams: false }));
server.listen(process.env.port || process.env.PORT || 3978, () => {
  console.log(`\nBot Started, ${server.name} listening to ${server.url}`);
});

// Listen for incoming bot request
server.post("/api/messages", async (req, res) => {
  await adapter.processActivity(req, res, async (context) => {
    await bot.run(context);
  });
});

// Notification HTTP trigger
server.post("/api/notification", async (req, res) => {
  const requestBody = req.body;

  if (!('receiver' in requestBody) || typeof (requestBody['receiver']) !== 'string') {
    res.json(400, { "error": "Invalid argument 'receiver'." });
    return;
  }

  if (!('content' in requestBody) || typeof (requestBody['content']) !== 'string') {
    res.json(400, { "error": "Invalid argument 'content'." });
    return;
  }

  const receiver: string = requestBody['receiver'];
  const content: string = requestBody['content'];

  const ref = await store.get();
  if (!ref) {
    res.json(400, { "error": "Cannot send notification before the app is installed to a team." });
    return;
  }

  try {
    const receiverConversationId = await getTeamMemberInfoByEmail(adapter, ref, receiver);
    if (!receiverConversationId) {
      res.json(400, { "error": "Invalid receiver email." });
      return;
    }

    // TODO: customize the message by passing custom content and construct an adaptive card.
    const message = MessageFactory.text(content);

    await notificationSender.sendActivityToMember(ref, receiverConversationId, message);
  } catch (error) {
    res.json(500, { "error": error });
  }

  res.send(204);
});

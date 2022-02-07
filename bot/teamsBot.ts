import { BotFrameworkAdapter, ChannelInfo, ConversationReference, TeamInfo, TeamsActivityHandler, TeamsChannelData, TeamsInfo, TurnContext } from "botbuilder";
import { ConversationReferenceStore } from "./store";

function cloneConversationReference(ref: Partial<ConversationReference>): Partial<ConversationReference> {
  return JSON.parse(JSON.stringify(ref));
}

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

export class TeamsBot extends TeamsActivityHandler {
  adapter: BotFrameworkAdapter;
  conversationReferenceStore: ConversationReferenceStore;
  botId: string;

  constructor(adapter: BotFrameworkAdapter, conversationReferenceStore: ConversationReferenceStore, botId: string) {
    super();
    this.adapter = adapter;
    this.conversationReferenceStore = conversationReferenceStore;
    this.botId = botId;

    this.onMembersAdded(async (context: TurnContext, next) => {
      let isSelfAdded: boolean = false;
      // check whether the member added is myself (the bot app)
      for (const member of context.activity.membersAdded) {
        if (member.id.includes(botId)) {
          isSelfAdded = true;
        }
      }

      const ref = TurnContext.getConversationReference(context.activity);
      // TODO: only support adding to a team channel
      if (isSelfAdded && context.activity.conversation.conversationType === "channel") {
        const channels = await TeamsInfo.getTeamChannels(context);
        channels.forEach((channel) => {
          if (channel.id) {
            const channelRef = cloneConversationReference(ref);
            channelRef.conversation.id = channel.id;
            channelRef.conversation.name = channel.name;
            this.conversationReferenceStore.add(channelRef);
          }
        });
      }

      await next();
    });

    this.onTeamsChannelCreatedEvent(async (channelInfo: ChannelInfo, teamInfo: TeamInfo, context: TurnContext, next) => {
      if (channelInfo.id) {
        const ref = TurnContext.getConversationReference(context.activity);
        const channelRef = cloneConversationReference(ref);
        channelRef.conversation.id = channelInfo.id;
        channelRef.conversation.name = channelInfo.name;
        this.conversationReferenceStore.add(channelRef);
      }

      await next();
    })

    // Set the onTurnError for the singleton BotFrameworkAdapter.
    this.adapter.onTurnError = onTurnErrorHandler;
  }
}
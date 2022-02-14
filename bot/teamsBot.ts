import { BotFrameworkAdapter, TeamInfo, TeamsActivityHandler, TurnContext } from "botbuilder";
import { ConversationReferenceStore } from "./sdk/conversationReferenceStore";

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

      if (isSelfAdded) {
        const ref = TurnContext.getConversationReference(context.activity);
        await this.conversationReferenceStore.add(ref);
      }

      await next();
    });

    this.onTeamsTeamDeletedEvent(async (teamInfo: TeamInfo, context: TurnContext, next) => {
      if (teamInfo.id) {
        await this.conversationReferenceStore.delete(teamInfo.id);
      }

      await next();
    })

    // Set the onTurnError for the singleton BotFrameworkAdapter.
    this.adapter.onTurnError = onTurnErrorHandler;
  }
}
import { BotFrameworkAdapter, TeamsActivityHandler, TurnContext } from "botbuilder";
import { ConversationReferenceStore } from "./store";

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

  constructor(adapter: BotFrameworkAdapter, conversationReferenceStore: ConversationReferenceStore) {
    super();
    this.adapter = adapter;
    this.conversationReferenceStore = conversationReferenceStore;

    this.onMembersAdded(async (context, next) => {
      this.conversationReferenceStore.set(TurnContext.getConversationReference(context.activity));
      await next();
    });

    // Set the onTurnError for the singleton BotFrameworkAdapter.
    this.adapter.onTurnError = onTurnErrorHandler;
  }
}
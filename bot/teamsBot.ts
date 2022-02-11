import { TeamsActivityHandler, CardFactory, TurnContext, MessageFactory, BotFrameworkAdapter } from "botbuilder";
import rawWelcomeCard from "./adaptiveCards/welcome.json"
import { AdaptiveCards } from "@microsoft/adaptivecards-tools";
import { ConversationReferenceStore } from "./sdk/store";
import { conversationIdToTeamId } from "./sdk/botUtils";

const ErrorMessages = {
  OnlySupportTeam: "This bot app only supports running in a Team, not group chat or personal chat",
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
  store: ConversationReferenceStore;
  botId: string;
  adapter: BotFrameworkAdapter;

  constructor(adapter: BotFrameworkAdapter, store: ConversationReferenceStore, botId: string) {
    super();

    // Set the onTurnError for the singleton BotFrameworkAdapter.
    adapter.onTurnError = onTurnErrorHandler;

    this.adapter = adapter;
    this.store = store;
    this.botId = botId;

    this.onMessage(async (context, next) => {
      if (context.activity.conversation.conversationType !== "channel") {
        await this.sendMessageActivity(context, ErrorMessages.OnlySupportTeam);
        await next();
        return;
      }

      let txt = context.activity.text;
      const removedMentionText = TurnContext.removeRecipientMention(
        context.activity
      );
      if (removedMentionText) {
        // Remove the line break
        txt = removedMentionText.toLowerCase().replace(/\n|\r/g, "").trim();
      }

      // Trigger command by IM text
      switch (txt) {
        case "welcome": {
          const card = AdaptiveCards.declareWithoutData(rawWelcomeCard).render();
          await context.sendActivity({ attachments: [CardFactory.adaptiveCard(card)] });
          break;
        }
        case "reset": {
          // reset conversation reference to current team in case conversation reference lost
          const ref = TurnContext.getConversationReference(context.activity);
          ref.conversation.id = conversationIdToTeamId(ref.conversation.id);
          await this.store.set(ref);
          break;
        }
      }

      await next();
    });

    this.onMembersAdded(async (context, next) => {
      let isSelfAdded = false;
      for (const member of context.activity.membersAdded) {
        if (member.id.includes(this.botId)) {
          isSelfAdded = true;
        }
      }
      
      if (isSelfAdded) {
        if (context.activity.conversation.conversationType !== "channel") {
          await this.sendMessageActivity(context, ErrorMessages.OnlySupportTeam);
          await next();
          return;
        }
        await this.store.set(TurnContext.getConversationReference(context.activity));
      }

      // Show welcome message when a member (including bot) is added to the channel.
      const card = AdaptiveCards.declareWithoutData(rawWelcomeCard).render();
      await context.sendActivity({ attachments: [CardFactory.adaptiveCard(card)] });

      await next();
    });
  }

  async sendMessageActivity(context: TurnContext, message: string): Promise<void> {
    await context.sendActivity(MessageFactory.text(message));
  }
}
import { BotFrameworkAdapter, TeamsActivityHandler, TurnContext } from "botbuilder";
import { ConversationReferenceStore } from "./store";

export class TeamsBot extends TeamsActivityHandler {
  adapter: BotFrameworkAdapter;
  conversationReferenceStore: ConversationReferenceStore;

  constructor(conversationReferenceStore: ConversationReferenceStore) {
    super();
    this.conversationReferenceStore = conversationReferenceStore;

    this.onMembersAdded(async (context, next) => {
      this.conversationReferenceStore.set(TurnContext.getConversationReference(context.activity));
      await next();
    });
  }
}
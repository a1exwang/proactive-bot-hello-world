import { BotFrameworkAdapter, ConversationReference, MessageFactory, TurnContext } from "botbuilder";

export class NotificationSender {
  adapter: BotFrameworkAdapter;
  converstationReference: Partial<ConversationReference>;

  constructor(adapter: BotFrameworkAdapter) {
    this.adapter = adapter;
  }

  public async sendNotification(conversationReference: Partial<ConversationReference>, text: string) {
    await this.adapter.continueConversation(conversationReference, async (context: TurnContext) => {
      const message = MessageFactory.text(text);
      await context.sendActivity(message);
    });
  }
}
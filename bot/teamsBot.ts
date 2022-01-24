import { BotFrameworkAdapter, ConversationReference, MessageFactory, TeamsActivityHandler, TurnContext } from "botbuilder";

export class TeamsBot extends TeamsActivityHandler {
  adapter: BotFrameworkAdapter;
  conversationReference: Partial<ConversationReference> | undefined

  constructor(adapter: BotFrameworkAdapter) {
    super();
    this.adapter = adapter;

    this.onMembersAdded(async (context, next) => {
      // store conversation reference when the bot app is added to a channel.
      // You can persist the conversationReference to a file, database, etc. to pro-actively send messages at any time.
      // To serialize conversation reference to JSON: 
      //    JSON.stringify(conversationReference);
      this.conversationReference = TurnContext.getConversationReference(context.activity);
      await next();
    });

    // trigger the notification on schedule.
    setInterval(() => {
      if (this.conversationReference !== undefined) {
        this.sendProactiveMessage(this.conversationReference)
      }
    }, 10000);
  }

  // send the proactive message to the conversation.
  async sendProactiveMessage(conversationReference: Partial<ConversationReference>) {
    const message = MessageFactory.text(`Hello world at ${new Date()}`);
    await this.adapter.continueConversation(conversationReference, async (context) => {
      await context.sendActivity(message);
    });
  }
}

